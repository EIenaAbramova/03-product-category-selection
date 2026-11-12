let
    T=Table.Buffer(Products),
    Rows=Table.RowCount(T),
    Unique=List.Count(List.Distinct(T[url])),
    Out=#table({"Показатель","Значение"},{
        {"Строк источника",Rows},
        {"Уникальных карточек",Unique},
        {"Повторных ссылок",Rows-Unique},
        {"Товаров свыше 1 млн",List.Sum(T[million])},
        {"Сумма заказов, ₽",List.Sum(T[revenue])},
        {"ДРР свыше 100%, сохранены",List.Count(List.Select(T[drr],each _>1))},
        {"ДРР равна нулю",List.Count(List.Select(T[drr],each _=0))},
        {"Не указан бренд",List.Count(List.Select(T[brand],each _=null or _=""))},
        {"Расходы расчетные, ₽",List.Sum(T[ad_spend_est])}
    })
in Out
