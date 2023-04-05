use [Global Trade]
SELECT * FROM Production

-- Working code
CREATE TABLE NewProduction (
  column1 nvarchar(50),
  column2 nvarchar(50),
  year int,
  value nvarchar(50)
);

INSERT INTO NewProduction (column1, column2, year, value)
SELECT column1, column2, '2011' AS year, column25 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2012' AS year, column26 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2013' AS year, column27 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2014' AS year, column28 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2015' AS year, column29 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2016' AS year, column30 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2017' AS year, column31 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2018' AS year, column32 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2019' AS year, column33 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2020' AS year, column34 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2021' AS year, column35 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2022' AS year, column36 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2023' AS year, column37 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2024' AS year, column38 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2025' AS year, column39 AS value
FROM Production
UNION ALL
SELECT column1, column2, '2026' AS year, column40 AS value
FROM Production
ORDER BY column1, column2, year;

-- Delete null columns
DELETE FROM NewProduction 
WHERE column1 IS NULL 
AND column2 IS NULL 
AND value IS NULL;

SELECT * FROM NewProduction;
