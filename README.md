# GoogleSheetsCRUD – a simple library to make basic CRUD operations in a Google Sheet

* C: create
* R: read
* U: update
* D: delete

## Usage

This library assumes that the first row of all manipulated ranges is the title of columns. 

In a new Google Sheets document, create a sheet named `People` with the following content:

```
id	name	country	city
1	Julie	France	Paris
2	Julien	France	Montpellier
3	Marek	Slovakia	Košice
4	Tobias	Austria	Vienna
5	Agnieszka	Poland	Sosnowiec
6	Giorgi	Georgia	Kutaisi
7	John	USA	Los Angeles
8	Ivan	Russia	Norilsk
9	Marina	Russia	Moscow
10	Andreas	Germany	Berlin
```

### Set up

```php
require('vendor/autoload.php');

use Mutusen\GoogleSheetsCRUD\GoogleSheetsCRUD;

$gs = new GoogleSheetsCRUD(
    'sheet id', // Found in the URL of the Google Sheet: https://docs.google.com/spreadsheets/d/.../edit
    'service account' // JSON object given by the Google Sheets API
);
```

### Read entire range
```php
/*
 * Returns:
 * Array
    (
        [0] => Array
            (
                [id] => 1
                [name] => Julie
                [country] => France
                [city] => Paris
            )
    
        [1] => Array
            (
                [id] => 2
                [name] => Julien
                [country] => France
                [city] => Montpellier
            )
        ...
    )
 */
$data = $gs->readAll('People');

// You can also specify a range in the sheet (works for all other functions except appendRow())
$data = $gs->readAll('People!B1:D6');
```

### Get a specific row

```php
/*
 * Returns:
 * Array
    (
        [id] => 1
        [name] => Julie
        [country] => France
        [city] => Paris
    )
 * If there are several matches, it stops at the first one
 */
$data = $gs->getRowWhere('People', 'id', 1);
```

### Get specific rows

```php
/*
 * Returns:
 * Array
    (
        [0] => Array
            (
                [id] => 1
                [name] => Julie
                [country] => France
                [city] => Paris
            )
    
        [1] => Array
            (
                [id] => 2
                [name] => Julien
                [country] => France
                [city] => Montpellier
            )
    
    )
 */
$data = $gs->getRowsWhere('People', 'country', 'France');
```

### Insert row

```php
// You cannot use a range after the name of the sheet
// The values have to be in the right order
$gs->appendRow('People', [11, 'Maria', 'Italy', 'Milan']);
```

### Insert multiple rows

```php
// You cannot use a range after the name of the sheet
// The values have to be in the right order
$rows = [
    [11, 'Maria', 'Italy', 'Milan'],
    [12, 'Oleh', 'Ukraine', 'Lviv']
];
$gs->appendRows('People', $rows);
```

### Update row

```php
// You can update multiple values in a single row
$gs->updateFieldsWhere('People', 'id', 11, [
    'country' => 'Spain',
    'city' => 'Madrid',
]);

// If you use a search criterion that matches several rows, they all will be updated
$gs->updateFieldsWhere('People', 'country', 'France', [
    'country' => 'Belgium',
    'city' => 'Brussels',
]);
```

### Update multiple rows

Calling `updateFieldsWhere()` several times in a row is inefficient because it would make an API call each time, instead use the following method:

```php
// If you want to update a sheet according to several criteria (e.g. several ids)
use Mutusen\GoogleSheetsCRUD\GSCMultipleLinesUpdate;

$query = new GSCMultipleLinesUpdate($gs, 'People');
$query->updateWhere('id', 9, ['name' => 'Irina']);
$query->updateWhere('id', 10, ['name' => 'Jonas']);
$query->updateWhere('id', 11, ['name' => 'Ines']);
$query->execute();
```

### Delete row

```php
$data = $gs->deleteRowWhere('People', 'id', 11);
```