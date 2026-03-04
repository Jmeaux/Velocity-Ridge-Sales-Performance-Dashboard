;WITH MaxDate AS
(
    SELECT MAX(CAST(OrderDate AS date)) AS MaxOrderDate
    FROM Sales.SalesOrderHeader
),
Params AS
(
    SELECT
        DATEADD(year, -2, MaxOrderDate) AS StartDate,
        MaxOrderDate                    AS EndDate
    FROM MaxDate
),
SalesLines AS
(
    SELECT
        CAST(h.OrderDate AS date) AS OrderDate,
        h.SalesOrderID,
        h.TerritoryID,
        d.ProductID,
        d.OrderQty,
        d.UnitPrice,
        d.UnitPriceDiscount,
        d.LineTotal
    FROM Sales.SalesOrderHeader h
    INNER JOIN Sales.SalesOrderDetail d
        ON d.SalesOrderID = h.SalesOrderID
    CROSS JOIN Params p
    WHERE
        h.OrderDate >= p.StartDate
        AND h.OrderDate < DATEADD(day, 1, p.EndDate)
        -- Optional shipped-only filter (uncomment if desired):
        -- AND h.Status = 5
),
Dimmed AS
(
    SELECT
        sl.OrderDate,
        sl.SalesOrderID,
        sl.TerritoryID,
        st.Name    AS TerritoryName,
        st.[Group] AS TerritoryGroup,

        pc.ProductCategoryID,
        pc.Name AS ProductCategory,
        psc.ProductSubcategoryID,
        psc.Name AS ProductSubcategory,

        sl.OrderQty,
        sl.LineTotal,
        CAST(sl.UnitPrice * sl.OrderQty * sl.UnitPriceDiscount AS money) AS DiscountAmount,
        CAST(pr.StandardCost * sl.OrderQty AS money) AS StandardCost
    FROM SalesLines sl
    LEFT JOIN Sales.SalesTerritory st
        ON st.TerritoryID = sl.TerritoryID
    INNER JOIN Production.Product pr
        ON pr.ProductID = sl.ProductID
    LEFT JOIN Production.ProductSubcategory psc
        ON psc.ProductSubcategoryID = pr.ProductSubcategoryID
    LEFT JOIN Production.ProductCategory pc
        ON pc.ProductCategoryID = psc.ProductCategoryID
)
SELECT
    d.OrderDate,

    d.TerritoryID,
    d.TerritoryName,
    d.TerritoryGroup,

    d.ProductCategoryID,
    d.ProductCategory,
    d.ProductSubcategoryID,
    d.ProductSubcategory,

    COUNT(DISTINCT d.SalesOrderID) AS OrderCount,
    SUM(d.OrderQty) AS Units,

    CAST(SUM(d.LineTotal) AS money) AS Revenue,
    CAST(SUM(d.DiscountAmount) AS money) AS DiscountAmount,
    CAST(SUM(d.StandardCost) AS money) AS StandardCost,
    CAST(SUM(d.LineTotal) - SUM(d.StandardCost) AS money) AS GrossProfit,

    CAST(
        CASE WHEN SUM(d.LineTotal) = 0 THEN 0
             ELSE (SUM(d.LineTotal) - SUM(d.StandardCost)) / SUM(d.LineTotal)
        END
    AS decimal(9,4)) AS GrossMarginPct
FROM Dimmed d
GROUP BY
    d.OrderDate,
    d.TerritoryID, d.TerritoryName, d.TerritoryGroup,
    d.ProductCategoryID, d.ProductCategory,
    d.ProductSubcategoryID, d.ProductSubcategory;
