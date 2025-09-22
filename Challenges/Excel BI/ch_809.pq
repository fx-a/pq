let
    // FUNCTIONS
    f = (r)=> [
        height = getHeight(r[String]),
        age = getAge(r[String]),
        dob = getDOB(r[String]),
        out = [Date of Birth = dob, Age = age, Height = height]
    ][out],

    getHeight = (txt)=> [
        spl = List.Buffer(Splitter.SplitTextByEachDelimiter({"'",""""})(txt)),
        height = if List.Count(spl)=3 then [
            foot = Text.End(spl{0},1) & "'",
            inch = spl{1} & """",
            out = foot & inch][out]
            else null
    ][height],

    getAge = (txt)=> [
        spl_yy = Text.Split(txt, " years old"),
        spl_at = Text.Split(txt, "at just "),
        age = if List.Count(spl_yy)>1 then [
                float = Text.AfterDelimiter(spl_yy{0}, " ", {0,1}),
                out = Number.From(Text.BeforeDelimiter(float,"."))
            ][out] else 
            if List.Count(spl_at)>1 then [
                float = Text.SplitAny(spl_at{1}, " ."){0},
                out = Number.From(float)                
            ][out] else
        null
    ][age],

    getDOB = (txt)=> [
        flags = {"born on ", "birth is ", "birthday on "},
        pos = List.RemoveItems(
            List.Transform(flags, each Text.PositionOf(txt, _, 0)),
            {-1}){0}? ??Text.Length(txt),
        spl = Splitter.SplitTextByPositions({pos})(txt),
        rem = Text.SplitAny(spl{0},",."){0},
        acc = List.Accumulate(flags, rem, (s,c)=> Text.Replace(s,c,"")),
        clean = [
            a=Text.Split(acc," "), 
            b=Text.Select(a{0},{"0".."9","/","."}), 
            c=Text.Combine({b} & List.Skip(a))][c],
        dob = Date.From(clean,"en-US")
    ][dob],

    // TRANSFORMATION STEPS
    src = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    add = Table.AddColumn(src, "A", f),
    res = Table.ExpandRecordColumn(add, "A", {"Date of Birth","Age","Height"})
in
    res