import unittest
from unittest.mock import patch, MagicMock
import datetime

# Mock GSheetsDB before importing app to avoid needing credentials.json
with patch('core.gsheets_db.GSheetsDB.__init__', return_value=None):
    from app import app
    import app as app_module

class TestAppRoutes(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        app.config['SECRET_KEY'] = 'test_secret'
        self.client = app.test_client()

        # Provide a mocked DB and Calculator on the app module
        self.mock_db = MagicMock()
        self.mock_calc = MagicMock()
        app_module.db = self.mock_db
        app_module.calc = self.mock_calc

    def test_index_route(self):
        """GET / should render the index page containing commodity and division inputs."""
        response = self.client.get('/')
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("GM Mini", html)
        self.assertIn("CEMT", html)
        self.assertIn("COAL", html)
        self.assertIn("ADI", html)

    def test_check_date_missing_date(self):
        """POST /api/check_date with no date returns error."""
        response = self.client.post('/api/check_date', data={})
        self.assertEqual(response.status_code, 200)
        json_data = response.get_json()
        self.assertEqual(json_data['status'], 'error')
        self.assertIn("No date provided", json_data['message'])

    def test_check_date_invalid_format(self):
        """POST /api/check_date with invalid format returns error."""
        response = self.client.post('/api/check_date', data={'report_date': '01-04-2026'})
        self.assertEqual(response.status_code, 200)
        json_data = response.get_json()
        self.assertEqual(json_data['status'], 'error')
        self.assertIn("Invalid date format", json_data['message'])

    def test_check_date_empty_db_must_start_april_1st(self):
        """When DB is empty for a new FY, the first date MUST be April 1st."""
        self.mock_db.get_last_filled_date.return_value = None

        # Trying to submit May 5th on empty FY 2026-27
        response = self.client.post('/api/check_date', data={'report_date': '2026-05-05'})
        json_data = response.get_json()
        self.assertEqual(json_data['status'], 'error')
        self.assertIn("You must start by filling data for 01-04-2026", json_data['message'])

        # Submitting April 1st on empty DB is valid
        response_valid = self.client.post('/api/check_date', data={'report_date': '2026-04-01'})
        json_valid = response_valid.get_json()
        self.assertEqual(json_valid['status'], 'ok')

    def test_check_date_chronological_skipping(self):
        """Cannot skip dates (e.g., last filled is April 1st, choosing April 3rd is an error)."""
        self.mock_db.get_last_filled_date.return_value = datetime.date(2026, 4, 1)

        response = self.client.post('/api/check_date', data={'report_date': '2026-04-03'})
        json_data = response.get_json()
        self.assertEqual(json_data['status'], 'error')
        self.assertIn("Cannot skip dates", json_data['message'])

    def test_check_date_already_filled_prompts_overwrite(self):
        """Choosing a date that was already filled returns a warning to confirm overwrite."""
        self.mock_db.get_last_filled_date.return_value = datetime.date(2026, 4, 5)

        response = self.client.post('/api/check_date', data={'report_date': '2026-04-02'})
        json_data = response.get_json()
        self.assertEqual(json_data['status'], 'warning')
        self.assertIn("already present in the database", json_data['message'])

    def test_check_date_next_sequential_date_ok(self):
        """Choosing next consecutive date returns ok."""
        self.mock_db.get_last_filled_date.return_value = datetime.date(2026, 4, 5)

        response = self.client.post('/api/check_date', data={'report_date': '2026-04-06'})
        json_data = response.get_json()
        self.assertEqual(json_data['status'], 'ok')

    @patch('app.upload_excel_to_drive')
    @patch('app.generate_daily_excel')
    def test_generate_report_success(self, mock_generate_excel, mock_upload_drive):
        """POST /generate with valid April 1st date runs pipeline and flashes success link."""
        self.mock_db.get_last_filled_date.return_value = None
        self.mock_calc.calculate_daily_report.return_value = {
            'commodity': {}, 'division': {}, 'summary': {}
        }
        mock_generate_excel.return_value = "generated_workbooks/01-04-2026.xlsx"
        mock_upload_drive.return_value = "https://drive.google.com/test-link"

        form_data = {
            'report_date': '2026-04-01',
            'CEMT_rakes': '9',
            'CEMT_wagons': '420',
            'ADI_rakes': '4',
            'ADI_wagons': '200'
        }
        response = self.client.post('/generate', data=form_data, follow_redirects=True)
        self.assertEqual(response.status_code, 200)

        # Verify DB save was invoked
        self.mock_db.save_daily_batch.assert_called_once()
        # Verify calculation was invoked
        self.mock_calc.calculate_daily_report.assert_called_once()
        # Verify Excel generation was invoked
        mock_generate_excel.assert_called_once()
        # Verify Drive upload was invoked
        mock_upload_drive.assert_called_once()

        # Verify success message rendered in response
        html = response.get_data(as_text=True)
        self.assertIn("Success! Excel report for 01-04-2026 has been generated", html)


if __name__ == '__main__':
    unittest.main()
