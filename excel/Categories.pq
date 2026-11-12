let
    Source = Table.Buffer(Products),
    Groups = Table.Group(Source, {"cat1","cat2","cat3","cat4"}, {
        {"Товаров",each Table.RowCount(_),Int64.Type},
        {"Более миллиона",each List.Sum([million]),Int64.Type},
        {"Сумма заказов",each List.Sum([revenue]),type number},
        {"Расходы расчетные",each List.Sum([ad_spend_est]),type number},
        {"ДРР средняя",each List.Average([drr]),type number},
        {"ДРР ноль",each List.Count(List.Select([drr],each _=0)),Int64.Type},
        {"Продавцов",each List.Count(List.Distinct([seller])),Int64.Type},
        {"Медианная цена",each List.Median([price]),type number},
        {"Лидер заказы",each List.Max(Table.Group(_, {"seller"}, {{"r",each List.Sum([revenue]),type number}})[r]),type number}
    }),
    Share = Table.AddColumn(Groups,"Доля миллионников",each [Более миллиона]/[Товаров],type number),
    DRR = Table.AddColumn(Share,"ДРР взвешенная",each [Расходы расчетные]/[Сумма заказов],type number),
    Leader = Table.AddColumn(DRR,"Доля лидера",each [Лидер заказы]/[Сумма заказов],type number),
    Names = Table.RenameColumns(Leader, {{"cat1","Уровень 1"},{"cat2","Уровень 2"},{"cat3","Уровень 3"},{"cat4","Категория"}}),
    Sorted = Table.Sort(Names, {{"Доля миллионников",Order.Descending},{"ДРР взвешенная",Order.Ascending},{"Товаров",Order.Descending}})
in
    Sorted
