SELECT
	s.ShipperName as Shipper, COUNT(distinct o.OrderID) as "Total Orders"
FROM
	Shippers s, orders o
WHERE
	s.ShipperID = O.ShipperID
	AND s.ShipperName = "Speedy Express"
;

SELECT 
	FirstName as "Employee First Name", LastName as "Employee Last Name", MAX(NumOrders) as Orders
FROM
	(SELECT
		e.FirstName, e.LastName, COUNT(distinct o.OrderID) as NumOrders
	FROM
		Employees e, orders o
	WHERE
		e.EmployeeID = o.EmployeeID
	GROUP BY e.EmployeeID
	)
;

SELECT
	Product, NumOrders as "Times Ordered", MAX(TotalQuantity) as "Total Quantity", 
	AvgQuantity as "AVG Quantity per Order"
FROM
	(SELECT
		p.ProductName as Product, count(od.productID) as NumOrders, 
		SUM(od.Quantity) as TotalQuantity, SUM(od.Quantity) / COUNT(od.productID) as AvgQuantity 
	FROM
		Customers c, Orders o, OrderDetails od, Products p
	WHERE
		c.CustomerID = o.CustomerID
		AND o.OrderID = od.OrderID
		AND od.ProductID = p.ProductID
		AND c.Country = "Germany"
	GROUP BY 1
	)
; 