import React, { useState } from "react";
import { useForm } from "react-hook-form";
import '../node_modules/bootstrap/dist/css/bootstrap.min.css'
import ExcelJS from 'exceljs';
import * as XLSX from "xlsx";
import moment from 'moment'
import { al } from "./amigologo";
import { SelectRulon } from "./selectrulon";
import { vseRulonColor } from "./vseRulonColor";
import parse from 'html-react-parser';
var FileSaver = require('file-saver');

type Inputs = {
  width: number,
  height: number,
  material: string,
  color: string
};

export default function App() {
  const { register, getValues, watch, formState: { errors } } = useForm();
  //const onSubmit: SubmitHandler<Inputs> = data => console.log(data);
  let text = 'АВЕНСИС'
  const mat = document.getElementById('mat') as HTMLSelectElement
  if (mat) {
    text = mat.options[mat.selectedIndex].text
  }

  let colorRulon = ''
  for (let i = 0; i < vseRulonColor.length; i++) {
    vseRulonColor[i][0] = vseRulonColor[i][0].toUpperCase().replaceAll('BO', 'BLACK-OUT').replaceAll('_', '')
  }
  //let kn = Date.now()
  let col = document.getElementById('col') as HTMLSelectElement

const [cl, setCl] = useState(parse('<option>300715-0225 белый</option><option>300715-1908 черный</option>'))

let clo = cl
  if (col) {
    clo = col.options[col.selectedIndex].text
  }

const onchan = () => {
  colorRulon = ''
    vseRulonColor.forEach((i) => {
      if (i[0] == text) {
        let c = i[1] + ' ' + i[2]
        colorRulon += '<option>' + c + '</option>'
      }
    })
    setCl(parse(colorRulon))
  }
 
  let w = watch("width")
  let h = watch("height")
  //let m = getValues("material")
  //let c = getValues("color")

  // тоже работает, но не форматирует
  let b: string | ArrayBuffer | null = ''

  const showFile = (e: React.FormEvent<HTMLInputElement>): void => {
    let files: FileList | null = e.currentTarget.files;
    let fil = files![0]
    console.log(fil);
    let reader = new FileReader();
    reader.readAsArrayBuffer(fil);
    reader.onload = async function () {
      b = reader.result
      const wb = XLSX.read(b, { cellStyles: true });
      let ws = wb.Sheets['UNI']
      XLSX.utils.sheet_add_aoa(ws, [[w],], { origin: "D8" })
      XLSX.utils.sheet_add_aoa(ws, [[h],], { origin: "E8" })
      XLSX.writeFile(wb, "new.xlsx", { cellStyles: true })
    };
    reader.onerror = function () {
      console.log(reader.error);
    };


  }

  const go = async () => {
    let date = 'Дата ' + moment().format("DD") + '.' + moment().format("MM") + '.' + moment().format("YYYY") + 'г.'
    let dt = moment().format("DD") + moment().format("MM") + moment().format("YY")
    const workbook = new ExcelJS.Workbook();
    const wsh = workbook.addWorksheet('UNI')
    wsh.getCell('A3').value = 'Название фирмы "ГерАрт"'
    wsh.getCell('A3').font = {
      name: 'Times New Roman',
      size: 16
    }
    wsh.getCell('C3').value = date
    wsh.getCell('C3').font = {
      name: 'Times New Roman',
      size: 16
    }
    const colu = [
      { name: 'Вид изделия' },
      { name: 'Наименование ткани' },
      { name: 'Цвет ткани' },
      { name: 'Ширина по ребру штапика UNI (м)' },
      { name: 'Высота по ребру штапика UNI (м)' },
      { name: 'Кол-во\n(шт.)' },
      { name: 'Управление\n(прав / лев)' },
      { name: 'Длина\nуправления\n(м)' },
      { name: 'Цвет\nфурнитуры' },
      { name: 'Со свер-\nлением' },
      { name: 'На скотч' },
      { name: 'Натяжитель\nцепи' },
    ]
    const massA = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    const widthCol = [23, 32, 27, 12, 12, 12, 12, 12, 12, 12, 12, 12]

    wsh.getRow(5).height = 155
    wsh.getRow(5).font = { name: 'Times New Roman', size: 11 };
    wsh.getRow(4).font = { name: 'Times New Roman', size: 11 };
    wsh.mergeCells('D4:E4');
    wsh.mergeCells('J4:K4');
    wsh.mergeCells('A4:A5');
    wsh.mergeCells('B4:B5');
    wsh.mergeCells('C4:C5');
    wsh.mergeCells('F4:F5');
    wsh.mergeCells('G4:G5');
    wsh.mergeCells('H4:H5');
    wsh.mergeCells('I4:I5');
    wsh.mergeCells('L4:L5');
    wsh.getCell('D4').value = 'UNI'
    wsh.getCell('D4').font = {
      name: 'Times New Roman',
      size: 14,
      bold: true
    };
    wsh.getCell('D4').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    wsh.getCell('D4').border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    }
    wsh.getCell('D4').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'dbeef4' },
    }
    wsh.getCell('J4').value = 'Тип установки'
    wsh.getCell('J4').font = {
      name: 'Times New Roman',
      size: 11
    };
    wsh.getCell('J4').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    wsh.getCell('J4').border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    }
    wsh.getCell('J4').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'dbeef4' },
    }


    for (let i = 0; i < 12; i++) {
      let col = massA[i]
      let c = col + 5
      let w = widthCol[i]
      wsh.getColumn(col).width = w
      wsh.getCell(c).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      wsh.getCell(c).border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' }
      }
      wsh.getCell(c).value = colu[i].name
      wsh.getCell(c).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'dbeef4' },
      }
    }

    wsh.getCell('A4').border = {
      left: { style: 'medium' },
      top: { style: 'medium' },
      bottom: { style: 'medium' },
    }
    wsh.getCell('L4').border = {
      right: { style: 'medium' },
      top: { style: 'medium' },
      bottom: { style: 'medium' },
    }

    //wsh.spliceRows(1, 0, [])
    //wsh.spliceRows(1, 0, [])
    wsh.getCell('A1').value = 'Бланк заказа на кассетные рулонные шторы UNI1, UNI2, UNI1-Зебра, UNI2-Зебра, UNI с пружиной'
    wsh.getCell('A1').alignment = { horizontal: 'right' }
    wsh.getCell('A1').font = { name: 'Times New Roman', size: 14, bold: true }
    wsh.mergeCells('A1:L1')

    wsh.spliceRows(6, 0, ['УНИ', text, clo, Math.ceil(w / 10) / 100, Math.ceil(h / 10) / 100, 1, 'прав', 'ст', 'бел', 'да'])
    wsh.getRow(6).alignment = { vertical: 'middle', horizontal: 'center' }
    wsh.getRow(6).font = { name: 'Times New Roman', size: 14 }
    for (let i = 0; i < 12; i++) {
      let col = massA[i]
      let c = col + 6
      wsh.getCell(c).border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    }
    wsh.getRow(6).height = 58
    wsh.getRow(7).height = 43

    wsh.getCell('D6').numFmt = '0.00'
    wsh.getCell('E6').numFmt = '0.00'

    wsh.getCell('A7').value = '   Подпись___________                             Печать__________                                 Оплату гарантируем_____________                    С техническими особенностями ознакомлены___________'

    const myBase64Image = "data:image/png;base64," + al
    const imageId2 = workbook.addImage({
      base64: myBase64Image,
      extension: 'png',
    })
    wsh.addImage(imageId2, 'J8:K11');

    const buffer = await workbook.xlsx.writeBuffer();
    FileSaver.saveAs(new Blob([buffer]), `Заявка_УНИ_${dt}.xlsx`)
  }


  //changeRulon()

  return (


    <>
      <form className="row" noValidate>

        <div className="container-fluid w-75 form-floating  form-control-sm">
          <input className="form-control" {...register("width", { required: true })} />
          <label>ширина, мм</label>
        </div>
        {errors.width && <span>введите ширину</span>}

        <div className="container-fluid w-75 form-floating  form-control-sm">
          <input className="form-control" {...register("height", { required: true })} />
          <label>высота, мм</label>
        </div>
        {errors.height && <span>введите высоту</span>}

        <div className="container-fluid w-75 form-floating  form-control-sm">
          <select id='mat' className="form-select" defaultValue='АВЕНСИС' {...register("material")} onClick={onchan} onChange={onchan}>
            {SelectRulon}
          </select>
          <label>материал</label>
        </div>
        {errors.material && <span>материал</span>}

        <div className="container-fluid w-75 form-floating  form-control-sm" id="torender">
          <select id='col' className="form-select" defaultValue='белый' {...register("color")} >
            {cl}
          </select>
          <label>цвет</label>
        </div>

        {errors.color && <span>цвет</span>}
 
        {/* <div className="d-flex justify-content-center mt-2">
          <input type="file" onChange={showFile}></input>
        </div> */}
      </form>
      <div className="d-flex justify-content-center mt-2">
        <button className="btn btn-outline-primary mt-3" onClick={go}>go</button>
      </div>
    </>

  )

}