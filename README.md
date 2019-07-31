# XLSX from SQL

For older versions of Microsoft Excel

For newer versions, read [this](https://docs.microsoft.com/en-us/azure/sql-database/sql-database-connect-excel)
## Setup
Estabilish a connection
<ul>
  <li><code>host</code>: The hostname of the database you are connecting to (Default: <code>localhost</code>)</li>
  <li><code>userS</code>: The MySQL user to authenticate as</li>
  <li><code>passwordS</code>: The password of that MySQL user</li>
  <li><code>databaseS</code>: Name of the database to use for this connection</li>
  <li><code>portS</code>: The port number to connect to</li>
</ul>

<code>table</code>: The output table's name

<code>query_count</code>: A count query to count the number of rows produced by the query (for the progress bar)

<code>query</code>: The SQL query

## Usage
To install dependencies

```cli
pip install requirements.txt
```

To run

```cli
python export.py
```
