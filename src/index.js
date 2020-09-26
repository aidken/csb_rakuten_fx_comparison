import "./styles.css";
import XLSX from "xlsx";
import papa from "papaparse";


let orders      = {};
let inventories = {};
let comparisons = {};

let updateComparison = function(orders, inventories, switchResult) {

  comparisons = {};

  let Comparison = function(order, inventory) {
    this.order = order;
    this.inventory = inventory;
    this.difference = function() {
      if (this.inventory < this.order) {
        return this.inventory - this.order;
      } else {
        return 0;
      }

    }
  }

  let allKeys = Object.keys(orders).concat(Object.keys(inventories));
  let allItemNumbers = Array.from(new Set(allKeys)).sort();

  allItemNumbers.forEach(function(itemNumber) {

    let qty_ordered = 0;
    let qty_inventory = 0;

    if (itemNumber in orders) {
      qty_ordered = orders[itemNumber];
    }

    if (itemNumber in inventories) {
      qty_inventory = inventories[itemNumber];
    }

    comparisons[itemNumber] = new Comparison(qty_ordered, qty_inventory);

  });

  let count = 0;
  let html1 = "";
  let html2 = "<p>注文数量と在庫数量との比較は次の通りです。</p><table><tr><th align=right>品目</th><th align=right>FX の在庫数量</th><th align=right>楽天での注文数量</th><th align=right>不足している数量</th></tr>";
  for (const key in comparisons) {
    if ( comparisons[key].difference() < 0 ) {
      count += 1;
    }
    html2 += `<tr><td align=right>${key}</td><td align=right>${comparisons[key].inventory}</td><td align=right>${comparisons[key].order}</td align=right><td align=right>${comparisons[key].difference()}</td></tr>`;
  }

  if (switchResult) {
    if ( count > 0) {
      html1 = `<p>FX の在庫が足りていな品目が ${count} 品目見つかりました。下の表に在庫数量と注文数量の比較が載っていますので、確認してください。</p>`;
    } else {
      html1 = "<p>FX の在庫はすべて足りています。</p>";
    }

  }
  html2 += "</table>";

  document.getElementById("output1").innerHTML = html1;
  document.getElementById("output2").innerHTML = html2;

}

let ExcelToJSON = function() {

  // initialize
  orders = {};

  // order class
  let Order = function(row) {
    this.itemNumber = row.B.toString();
    this.qty        = row.E;
  };

  this.parseExcel = function(file) {
    var reader = new FileReader();

    reader.onload = function(evt) {
      let data          = evt.target.result;
      let workbook      = XLSX.read(data, {type: 'binary'});
      let worksheet     = workbook.Sheets['提出用'];
      let XL_row_object = XLSX.utils.sheet_to_json(worksheet, {header: "A"});

      XL_row_object.forEach(function(i) {
        // console.log(i);
        // if (typeof i.B !== 'undefined' and i.B.toString() != '商品ｺｰﾄﾞ') {
        if (typeof i.B !== 'undefined' && i.B.toString() !== '商品ｺｰﾄﾞ' && i.B.toString() !== '総計') {
          let order = new Order(i);
          if (order.itemNumber in orders) {
            console.log(`Strange, this item number ${order.itemNumber} appears more than once.`);
            orders[order.itemNumber] += order.qty;
          } else {
            orders[order.itemNumber] = order.qty;
          }
        }  // end if
      }); // end XL_row_object.forEach

      updateComparison(orders, inventories, true);

    };  // end onload

    reader.onerror = function(ex) {
      console.log(ex);
    };  // end reader.onerror

    reader.readAsBinaryString(file);
  }; // end parseExcel

}; // end function ExcelToJSON

function handleFileSelect1(evt) {
  let files = evt.target.files; // FileList object
  let xl2json = new ExcelToJSON();
  xl2json.parseExcel(files[0]);
}; // end function handleFile1


function handleFileSelect2(evt) {

  // initialize
  inventories = {};

  let Inventory = function (record) {
    this.itemNumber = record[0].toString();
    this.warehouse = record[1];
    this.location = record[2].toString();
    this.qty = record[3];
    this.tenDigitLot = record[4].toString();
  };

  let file = evt.target.files[0];
  papa.parse(file, {
    header: false,
    dynamicTyping: true,
    complete: function (results) {
      results.data.forEach(function (d) {
        if (d.length === 5) {
          let x = new Inventory(d);
          if (x.warehouse === "FX") {
            if (x.itemNumber in inventories) {
              inventories[x.itemNumber] += x.qty;
            } else {
              inventories[x.itemNumber] = x.qty;
            }
          }
        }
      }); // end forEach

      updateComparison(orders, inventories, false);

    },
  });
}

document.getElementById("app").innerHTML = `
	<form enctype="multipart/form-data">
		<div>
		<label for='upload2'>在庫のテキストファイル ili.txt を指定してください。</label>
		<input id="upload2" type=file name="files2[]" accept='.txt'>
		</div>
		<div>
		<label for='upload1'>楽天の注文の Excel ファイルをアップロードして下さい。</label>
		<input id="upload1" type=file name="files1[]" accept='.xlsm, .xlsx'>
		</div>
	</form>
`;

document.getElementById("upload1").addEventListener("change", handleFileSelect1, false);
document.getElementById("upload2").addEventListener("change", handleFileSelect2, false);
