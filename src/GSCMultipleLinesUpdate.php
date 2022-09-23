<?php
namespace Mutusen\GoogleSheetsCRUD;

class GSCMultipleLinesUpdate
{
	/**
	 * @var GoogleSheetsCRUD
	 */
	private GoogleSheetsCRUD $gsc;

	/**
	 * @var string
	 */
	private string $range;

	/**
	 * @var array
	 */
	private array $sheetData;

	/**
	 * @var array
	 */
	private array $dataToUpdate = [];

	/**
	 * Sets up a query to update multiple rows
	 * @param GoogleSheetsCRUD $gsc
	 * @param string $range Name of sheet, optionally with the range you want to read (e.g. Sheet1!A1:D10)
	 */
	public function __construct(GoogleSheetsCRUD $gsc, string $range)
	{
		$this->gsc = $gsc;
		$this->range = $range;
		$this->sheetData = $gsc->readAll($range);
	}

	/**
	 * Adds a row to be updated to the query
	 * @param string $fieldName Name of the field whose value you want to compare
	 * @param mixed $fieldValue The value you want to find
	 * @param array $newValues An associative array of field => value for the data you want to update
	 * @throws GoogleSheetsCRUDException
	 */
	public function updateWhere(string $fieldName, mixed $fieldValue, array $newValues): void
	{
		$rowIds = $this->gsc->findRowIndicesWhere($this->sheetData, $fieldName, $fieldValue);
		if (count($rowIds) == 0) {
			return;
		}

		foreach ($rowIds as $rowId) {
			// It is the row number in Google Sheets (i.e. starting from 1)
			// +2 to account for the PHP array starting at 0 and the row with field names
			$rowId = $rowId + 2 + $this->gsc->numberOfRowsBeforeRange($this->range);

			$sheetName = $this->gsc->getSheetName($this->range);
			foreach ($newValues as $field => $value) {
				$column_letter = $this->gsc->getColumnLetter($field, $this->sheetData, $this->gsc->numberOfColumnsBeforeRange($this->range));
				if (empty($column_letter)) {
					continue;
				}
				$this->dataToUpdate[] = new \Google_Service_Sheets_ValueRange([
					'range' => $sheetName . '!' . $column_letter . $rowId,
					'values' => [[$value]]
				]);
			}
		}
	}

	/**
	 * Updates the rows in the Google sheet
	 * @param string $valueInputOption See https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
	 * @throws GoogleSheetsCRUDException
	 */
	public function execute(string $valueInputOption = 'RAW'): void
	{
		$body = new \Google_Service_Sheets_BatchUpdateValuesRequest([
			'valueInputOption' => $valueInputOption,
			'data' => $this->dataToUpdate
		]);

		try {
			$this->gsc->getSheetService()->spreadsheets_values->batchUpdate($this->gsc->getFileId(), $body);
		}
		catch (\Google_Service_Exception $e) {
			throw new GoogleSheetsCRUDException($e->getMessage());
		}
	}
}
