let
    // FUNCTIONS
    f = (name) as table => Excel.CurrentWorkbook(){[Name = name]}[Content],
    g = (txt) as text => List.Accumulate({"_",".","/"}, Text.Range(txt,5), (s,c)=> Text.Replace(s,c,"-")),
    h = (tbl, rectype) as table => Table.SelectRows(tbl, each [Record_Type] = rectype),
    k = (lst, prefix, skipend) as list => [
        trlst = List.RemoveLastN(lst, skipend),
        out = List.Transform(trlst, each prefix & _) & List.LastN(lst, skipend)
        ][out],

    // TRANSFORMATION STEPS
    orders = Table.TransformColumns(f("Orders"), {"Cust_Code", g}),
    nms = {"Order_ID","Order_Date","Customer_Name","Customer_Contact",
           "Customer_Region","Product_Name","Product_Category","Quantity",
           "Unit_Price", "Order_Value", "Order_Status"},
    cust = Table.Buffer(f("Customers")),
    
    customers = h(cust, "CUSTOMER"),
    joinC = Table.NestedJoin(orders, "Cust_Code", customers, "ID_Code", "C", JoinKind.LeftOuter),
    expC = [nms = {"Name","Contact","Region"}, prefix = "Customer_", 
            out = Table.ExpandTableColumn(joinC, "C", nms, k(nms, prefix, 0))][out],

    products = h(cust, "PRODUCT"),
    joinP = Table.NestedJoin(expC, "Product_SKU", products, "ID_Code", "P", JoinKind.LeftOuter),
    expP = [nms = {"Name","Category","Unit_Price"}, prefix = "Product_", 
            out = Table.ExpandTableColumn(joinP, "P", nms, k(nms, prefix, 1))][out],
            
    add = Table.AddColumn(expP, "Order_Value", each [Qty] * [Unit_Price]),
    rnm = Table.RenameColumns(add, {{"Qty", "Quantity"},{"Status", "Order_Status"}}),
    reod = Table.ReorderColumns(rnm, nms),
    res = Table.Sort(reod, {"Order_Date", 0})
in
    res
