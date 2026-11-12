let
    P = Excel.CurrentWorkbook(){[Name="Settings"]}[Content],
    MaxDRR = Number.From(P{0}[ДРР]),
    MinShare = Number.From(P{0}[Доля]),
    MinCount = Number.From(P{0}[Товаров]),
    Chosen = Table.SelectRows(Categories, each [Товаров]>=MinCount and [ДРР взвешенная]<=MaxDRR and [Доля миллионников]>=MinShare),
    Joined = Table.NestedJoin(Products,{"cat1","cat2","cat3","cat4"},Chosen,{"Уровень 1","Уровень 2","Уровень 3","Категория"},"Category",JoinKind.Inner),
    Only = Table.SelectRows(Joined,each [million]=1 and [drr]<=MaxDRR),
    Groups = Table.Group(Only,{"cat1","cat2","cat3","cat4"},{{"Examples",each Table.FirstN(Table.Sort(_,{{"drr",Order.Ascending},{"revenue",Order.Descending}}),3)}}),
    Expand = Table.ExpandTableColumn(Groups,"Examples",{"name","seller","revenue","drr","price","source_row","url","scheme"}),
    Names = Table.RenameColumns(Expand,{{"cat1","Уровень 1"},{"cat2","Уровень 2"},{"cat3","Уровень 3"},{"cat4","Категория"},{"name","Товар"},{"seller","Продавец"},{"revenue","Заказы, ₽"},{"drr","ДРР"},{"price","Цена, ₽"},{"source_row","Строка источника"},{"url","Карточка"},{"scheme","Схема"}})
in Names
