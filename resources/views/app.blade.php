<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <title>Trio Project</title>
    <style>
        a {
            text-decoration: none
        }

        @font-face {
            font-weight: 400;
            font-style: normal;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-Regular.woff2') format('woff2');
        }

        @font-face {
            font-weight: 400;
            font-style: italic;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-Italic.woff2') format('woff2');
        }

        @font-face {
            font-weight: 500;
            font-style: normal;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-Medium.woff2') format('woff2');
        }

        @font-face {
            font-weight: 500;
            font-style: italic;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-MediumItalic.woff2') format('woff2');
        }

        @font-face {
            font-weight: 700;
            font-style: normal;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-Bold.woff2') format('woff2');
        }

        @font-face {
            font-weight: 700;
            font-style: italic;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-BoldItalic.woff2') format('woff2');
        }

        @font-face {
            font-weight: 900;
            font-style: normal;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-Black.woff2') format('woff2');
        }

        @font-face {
            font-weight: 900;
            font-style: italic;
            font-family: 'Inter-Loom';
            src: url('https://cdn.loom.com/assets/fonts/inter/Inter-UI-BlackItalic.woff2') format('woff2');
        }

        ul {
            list-style-type: decimal;
        }

        li.done, p.done {
            text-decoration: line-through;
        }
    </style>
</head>
<body>
<div>
    RR rate
    <input type="text" name="rate" id="rate">
</div>
<br>
<input type="file" name="xlfile" id="xlf">

<input type="checkbox" name="useworker" checked="" hidden>
<input type="checkbox" name="userabs" checked="" hidden>
<p id="materialCode" style="margin-bottom: 20px;">Trio_Sl.xlsx</p>
<ul>
    <li id="items">ProductsRems.xlsx (трих код)</li>
    <li id="initialPrice">PricesChange (2).xlsx (себестоимость)</li>
    <li id="price">PricesChange.xlsx (розничная цена)</li>
    <li id="partnerCode">Customers.xlsx</li>
    <li id="import">Импорт OOO__Trio_Prodjekt__TRIO_PRODJ_*******</li>
    <li id="return">ProductsOpsByCust.xlsx</li>
    <li id="salesAnalyse">SalesAnalyse.xlsx</li>
    <li id="sales">Sales</li>
</ul>
<div>
    <button onclick="createCSV()" id="createCSV">Создать CSV на день</button>
</div>
<div style="margin-top: 20px;">
    <button onclick="startNewDay()" id="newDay">Начать новый день</button>
</div>
<hr/>
<div>
    <button onclick="finalCreateCSV()" id="finalCreateCSV">Создать итоговый CSV</button>
</div>
<div id="days"></div>

<script
    src="https://code.jquery.com/jquery-3.4.1.min.js"
    integrity="sha256-CSXorXvZcTkaix6Yvo6HppcZGetbYMGWSFlBw8HfCJo="
    crossorigin="anonymous"></script>
<script src="/assets/js/xlsx.full.min.js"></script>
<script src="/assets/js/constants.js"></script>
<script src="/assets/js/index.js"></script>
</body>
</html>
