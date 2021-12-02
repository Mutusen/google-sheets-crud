<?php
declare(strict_types=1);

namespace Mutusen\GoogleSheetsCRUD;

require(__DIR__ . '/GSCMultipleLinesUpdate.php');

class GoogleSheetsCRUDException extends \Exception {}

class GoogleSheetsCRUD
{
	/**
	 * @var string
	 */
	private string $fileId;

	/**
	 * @var string
	 */
	private string $serviceAccount;

	/**
	 * @var \Google_Service_Sheets
	 */
	private \Google_Service_Sheets $sheetService;

	/**
	 * @param string $fileId ID of the file (found in the URL of the Google sheet)
	 * @param string $serviceAccount JSON given by the Google Sheets API
	 * @throws \Google\Exception
	 */
	public function __construct(string $fileId, string $serviceAccount)
	{
		$this->fileId = $fileId;
		$this->serviceAccount = $serviceAccount;

		if (!defined('SCOPES')) {
			define('SCOPES', implode(' ', array(
					\Google_Service_Sheets::SPREADSHEETS)
			));
		}

		$client = new \Google_Client();
		$client->setAuthConfig(json_decode($this->serviceAccount, true));
		$client->setScopes(SCOPES);
		$this->sheetService = new \Google_Service_Sheets($client);
	}

	/**
	 * @return string
	 */
	public function getFileId(): string
	{
		return $this->fileId;
	}

	/**
	 * @return \Google_Service_Sheets
	 */
	public function getSheetService(): \Google_Service_Sheets
	{
		return $this->sheetService;
	}

	/**
	 * Fetches a range from Google sheet document
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 * @return array
	 */
	private function getWholeRange(string $range): array
	{
		$response = $this->sheetService->spreadsheets_values->get($this->fileId, $range);
		$values = $response->getValues();
		return $values;
	}

	/**
	 * Takes an array such as:
	 * title1	title2
	 * val11	val12
	 * val21	val22

	 * and returns an array such as
	 * [
	 *  [title1 => val11, title2 => val12]
	 *  [title1 => val21, title2 => val22]
	 * ]
	 * @param array $values
	 * @return array
	 */
	private function arrayFromColumnTitles(array $values): array
	{
		$titles = [];
		for ($i = 0; $i < count($values[0]); $i++) {
			if ($values[0][$i] != '') {
				$titles[$i] = $values[0][$i];
			}
			else {
				$titles[$i] = $i;
			}
		}
		$last_column = $i;
		unset($values[0]); // Remove the row with titles
		$array = [];

		foreach ($values as $values_row) {
			$new_row = [];
			for ($i = 0; $i <= $last_column; $i++) {
				if (isset($titles[$i])) {
					if (isset($values_row[$i])) {
						$new_row[$titles[$i]] = $values_row[$i];
					}
					else {
						$new_row[$titles[$i]] = '';
					}
				}
			}
			$array[] = $new_row;
		}
		return $array;
	}

	/**
	 * Takes a name of range, e.g. Sheet1!A1:D10, and returns the number of lines before the range in the spreadsheet (in this case 0).
	 * Examples:
	 *  Sheet1: 0
	 *  Sheet1!A:E: 0
	 *  Sheet1!B4:C11: 3
	 * @param string $range Name of sheet, optionally with the range you want to read
	 * @return int
	 */
	public function numberOfRowsBeforeRange(string $range): int
	{
		$matches = [];
		if (!preg_match('#![A-Z]+([0-9]*):[A-Z]+([0-9]*)$#', $range, $matches)) {
			return 0;
		}

		if (empty($matches[1])) {
			$matches[1] = 1;
		}
		if (empty($matches[2])) {
			$matches[2] = 1;
		}
		return min($matches[1], $matches[2]) - 1;
	}

	public function numberOfColumnsBeforeRange(string $range): int
	{
		$matches = [];
		if (!preg_match('#!([A-Z]+)[0-9]*:([A-Z]+)[0-9]*$#', $range, $matches)) {
			return 0;
		}

		if (empty($matches[1])) {
			$matches[1] = 'A';
		}
		if (empty($matches[2])) {
			$matches[2] = 'A';
		}

		$column1 = $this->alpha2num($matches[1]);
		$column2 = $this->alpha2num($matches[2]);

		return min($column1, $column2);
	}

	/**
	 * Removes the range at the end of a sheet name
	 * @param string $range
	 * @return string
	 */
	public function getSheetName(string $range): string
	{
		return preg_replace('#![A-Z]+([0-9]*):[A-Z]+([0-9]*)$#', '', $range);
	}

	/**
	 * Converts an alphabetic string into an integer.
	 *
	 * @param int $n This is the number to convert.
	 * @return string The converted number.
	 * @author Theriault
	 *
	 */
	private function alpha2num($a) {
		$r = 0;
		$l = strlen($a);
		for ($i = 0; $i < $l; $i++) {
			$r += pow(26, $i) * (ord($a[$l - $i - 1]) - 0x40);
		}
		return $r - 1;
	}

	/**
	 * Converts an integer into the alphabet base (A-Z). Note that A = 0.
	 * https://www.php.net/manual/en/function.base-convert.php#94874
	 *
	 * @param int $n This is the number to convert.
	 * @return string The converted number.
	 * @author Theriault
	 *
	 */
	private function num2alpha(int $n): string
	{
		$r = '';
		for ($i = 1; $n >= 0 && $i < 10; $i++) {
			$r = chr(0x41 + ($n % pow(26, $i) / pow(26, $i - 1))) . $r;
			$n -= pow(26, $i);
		}
		return $r;
	}

	/**
	 * @param string $columnName
	 * @param array $sheetData
	 * @param int $columnsBefore
	 * @return string
	 */
	public function getColumnLetter(string $columnName, array $sheetData, int $columnsBefore): string
	{
		$columnNumber = 0;
		$keys = array_keys($sheetData[0]);
		for ($i = 0; $i < count($keys); $i++) {
			if ($keys[$i] == $columnName) {
				$columnNumber = $i;
			}
		}

		return $this->num2alpha($columnNumber + $columnsBefore);
	}

	/**
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 * @param bool $hasHeader Is the first row of the range the name of fields?
	 * @return array
	 */
	public function readAll(string $range, bool $hasHeader = true): array
	{
		$sheetData = $this->getWholeRange($range);

		if (!$hasHeader) {
			return $sheetData;
		}
		return $this->arrayFromColumnTitles($sheetData);
	}

	/**
	 * Returns the index of the first row where a field has a specific value. Returns false if nothing is found.
	 * @param array $sheetData
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @return int|false
	 * @throws GoogleSheetsCRUDException
	 */
	public function findRowIndexWhere(array $sheetData, string $fieldName, mixed $fieldValue): int|false
	{
		$i = 0;
		foreach ($sheetData as $row) {
			if (!isset($row[$fieldName])) {
				throw new GoogleSheetsCRUDException('There is no field named ' . htmlspecialchars($fieldName));
			}

			if ($row[$fieldName] == $fieldValue) {
				return $i;
			}
			$i++;
		}

		return false;
	}

	/**
	 * Returns the indexes of all rows where a field has a specific value. Returns an empty array if nothing is found.
	 * @param array $sheetData
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @return array
	 * @throws GoogleSheetsCRUDException
	 */
	public function findRowIndicesWhere(array $sheetData, string $fieldName, mixed $fieldValue): array
	{
		$found = [];
		$i = 0;
		foreach ($sheetData as $row) {
			if (!isset($row[$fieldName])) {
				throw new GoogleSheetsCRUDException('There is no field named ' . htmlspecialchars($fieldName));
			}

			if ($row[$fieldName] == $fieldValue) {
				$found[] = $i;
			}
			$i++;
		}

		return $found;
	}

	/**
	 * Returns the first row where a field has a specific value. Returns false if nothing is found.
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @return array|false
	 * @throws GoogleSheetsCRUDException
	 */
	public function getRowWhere(string $range, string $fieldName, mixed $fieldValue): array|false
	{
		$sheetData = $this->readAll($range);
		$rowId = $this->findRowIndexWhere($sheetData, $fieldName, $fieldValue);
		if ($rowId === false) {
			return false;
		}
		return $sheetData[$rowId];
	}

	/**
	 * Returns all rows where a field has a specific value. Returns an empty array if nothing is found.
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @return array
	 * @throws GoogleSheetsCRUDException
	 */
	public function getRowsWhere(string $range, string $fieldName, mixed $fieldValue): array
	{
		$sheetData = $this->readAll($range);
		$rowIds = $this->findRowIndicesWhere($sheetData, $fieldName, $fieldValue);

		$rows = [];

		foreach ($rowIds as $id) {
			$rows[] = $sheetData[$id];
		}
		return $rows;
	}

	/**
	 * Updates a value or a set of values in a row found by getRowWhere()
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @param array $newValues An associative array of field => value for the data you want to update
	 * @param string $valueInputOption See https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
	 * @throws GoogleSheetsCRUDException
	 */
	public function updateFieldsWhere(string $range, string $fieldName, mixed $fieldValue, array $newValues, string $valueInputOption = 'RAW'): void
	{
		$query = new GSCMultipleLinesUpdate($this, $range);
		$query->updateWhere($fieldName, $fieldValue, $newValues);
		$query->execute($valueInputOption);
	}

	/**
	 * Appends row at the bottom of a sheet
	 * @param string $sheet Name of sheet. You cannot use a specific range in a sheet.
	 * @param array $values
	 * @param string $valueInputOption See https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
	 * @throws GoogleSheetsCRUDException
	 */
	public function appendRow(string $sheet, array $values, string $valueInputOption = 'RAW'): void
	{
		$conf = ['valueInputOption' => 'RAW'];

		$requestBody = new \Google_Service_Sheets_ValueRange();
		$requestBody->setValues(['values' => array_values($values)]);

		try {
			$response = $this->sheetService->spreadsheets_values->append($this->fileId, $sheet, $requestBody, $conf);
		}
		catch (\Google_Service_Exception $e) {
			throw new GoogleSheetsCRUDException($e->getMessage());
		}
	}

	/**
	 * Finds the sheet ID from the name of the sheet
	 * @param string $name
	 * @return int	-1 if there is no sheet with the name
	 */
	private function findSheetId(string $name): int
	{
		$response = $this->sheetService->spreadsheets->get($this->fileId);
		$sheets = $response->getSheets();

		foreach ($sheets as $sheet) {
			$properties = $sheet->getProperties();
			if ($properties->getTitle() == $name) {
				return $properties->getSheetId();
			}
		}

		return -1;
	}

	/**
	 * Deletes the first row where a field has a specific value.
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @throws GoogleSheetsCRUDException
	 */
	public function deleteRowWhere(string $range, string $fieldName, mixed $fieldValue): void
	{
		$sheetData = $this->readAll($range);
		$index = $this->findRowIndexWhere($sheetData, $fieldName, $fieldValue);
		if ($index !== false) {
			$index += $this->numberOfRowsBeforeRange($range) + 1;

			$sheetId = $this->findSheetId($this->getSheetName($range));

			if ($sheetId == -1) {
				return;
			}

			$request = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
				'requests' => [
					'deleteDimension' => [
						'range' => [
							'sheetId' => $sheetId,
							'dimension' => 'ROWS',
							'startIndex' => $index,
							'endIndex' => $index + 1
						]
					]
				]
			]);

			$this->sheetService->spreadsheets->batchUpdate($this->getFileId(), $request);
		}
	}
}