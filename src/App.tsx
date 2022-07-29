import React, { useEffect, useState } from "react";
import { useForm } from "react-hook-form";
import '../node_modules/bootstrap/dist/css/bootstrap.min.css'
import ExcelJS from 'exceljs';
import * as XLSX from "xlsx";
import moment, { min } from 'moment'
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

  const clck = new Event("click")

  const { register, getValues, watch, formState: { errors }, handleSubmit } = useForm(
    {
      //mode: "onBlur"
      mode: "onChange"
    }
  );
  //const onSubmit: SubmitHandler<Inputs> = data => console.log(data);
  //let text = 'АВЕНСИС'
  const [text, setText] = useState('АВЕНСИС')
  const mat = document.getElementById('mat') as HTMLSelectElement


  let colorRulon = ''
  useEffect(() => {
    for (let i = 0; i < vseRulonColor.length; i++) {
      vseRulonColor[i][0] = vseRulonColor[i][0].toUpperCase().replaceAll('BO', 'BLACK-OUT').replaceAll('_', '')
    }
  })
  //let kn = Date.now()
  const col = document.getElementById('col') as HTMLSelectElement

  const [cl, setCl] = useState(parse('<option>300715-0225 белый</option><option>300715-1908 черный</option>'))

  const [clo, setClo] = useState('300715-0225 белый')
  //let clo = cl
  const onClo = () => {
    if (col) {
      setClo(col.options[col.selectedIndex].text)
    }
  }

  const onText = () => {
    if (mat) {
      setText(mat.options[mat.selectedIndex].text)
    }
  }

  const onchan = async () => {
    if (mat) {
      setText(mat.options[mat.selectedIndex].text)
    }
    colorRulon = ''

    let promise = new Promise((resolve, reject) => {
      vseRulonColor.forEach((i) => {
        if (i[0] == text) {
          let c = i[1] + ' ' + i[2]
          colorRulon += '<option>' + c + '</option>'
          //resolve(colorRulon += '<option>' + c + '</option>')
        }
      })

      resolve(setCl(parse(colorRulon)))
    })
    //.then(value => value)
    await promise
    //setCl(parse(colorRulon))
    //setClo(col.options[col.selectedIndex].text)
    //console.log(col.childNodes[0].textContent)
    //onClo()
    //ust()
    filter()
  }

  const onchangeFilter = async () => {
    /* onText()
    colorRulon = ''
    ustanovit_spisok_color()
    setCl(parse(colorRulon)) */


    //let ust_color = (n[0][1] + ' ' + n[0][2])
    //setClo(ust_color)

    /* const click = () => {
      if (mat)mat.dispatchEvent(clck)
      console.log(clo)
    }
    
        setTimeout(click, 100) */
    

    new Promise(resolve => resolve(onText()))
    .then(() => colorRulon = '')
    .then(() => ustanovit_spisok_color())
    .then(() => setCl(parse(colorRulon)))
    .then(() => setCl(parse(colorRulon)))
    .then(() => {if (mat)mat.dispatchEvent(clck)})
    .then(() => onClo())

  }

  const ustanovit_spisok_color = () => {
    const n = vseRulonColor.filter(item => item[0] == text)
    n.forEach((i) => {
      const c = (i[1] + ' ' + i[2])
      colorRulon += '<option>' + c + '</option>'
    })
    return colorRulon
  }

  useEffect(() => {
    //setCl(parse(colorRulon))
    //parse(colorRulon)
    ustanovit_spisok_color()
  })

  /* const ust = () => {
    const n = vseRulonColor.find(item => item[0] == text)
    if (n) setClo(n[1] + ' ' + n[2])
    //console.log(n)
    //console.log(clo)
  } */

  const filter = () => {
    const f = vseRulonColor.filter(item => item[0] == text)
    const f0 = (f[0][1] + ' ' + f[0][2])
    setClo(f0)
  }

  /* useEffect(() => {
    //ust()
    filter()
  }) */

  const [u, setU] = useState('прав')
  const upr = document.getElementById('upr') as HTMLSelectElement
  const onchanupr = () => {
    if (upr) {
      setU(upr.options[upr.selectedIndex].text)
    }
  }
  /* useEffect(() => {
    if (upr) u = upr.options[upr.selectedIndex].text
    if (col) clo = col.options[col.selectedIndex].text
    if (mat) text = mat.options[mat.selectedIndex].text
  }); */

  let w: number, h: number, num: number, stroka: any[], tabl: any[], tablstr: string
  w = watch("width")
  h = watch("height")
  num = watch("num")
  tabl = []

  const [arrtabl, setArrtabl] = useState(tabl)
  const [spisok, setSpisok] = useState('')

  //tablstr = ''
  //u = getValues("upr")
  //console.log(w,h,num,u)
  //let u = getValues("upr")
  stroka = ['УНИ-1', text, clo, Math.ceil(w) / 1000, Math.ceil(h) / 1000, num, u, 'ст', 'бел', 'да']




  // тоже работает, но не форматирует
  /* let b: string | ArrayBuffer | null = ''
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
  } */

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

    for (let i = 0; i < lengthSpisok; i++) {
      let rw = 6 + i
      wsh.spliceRows(rw, 0, arrtabl[i])
      wsh.getRow(rw).alignment = { vertical: 'middle', horizontal: 'center' }
      wsh.getRow(rw).font = { name: 'Times New Roman', size: 14 }
      for (let j = 0; j < 12; j++) {
        let col = massA[j]
        let c = col + rw
        wsh.getCell(c).border = { left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
      }
      wsh.getRow(rw).height = 58
      wsh.getRow(rw + 1).height = 43

      wsh.getCell('D' + rw).numFmt = '0.000'
      wsh.getCell('E' + rw).numFmt = '0.000'

      wsh.getCell('A' + (rw + 1)).value = '   Подпись___________                             Печать__________                                 Оплату гарантируем_____________                    С техническими особенностями ознакомлены___________'

    }
    const myBase64Image = "data:image/png;base64," + al
    const imageId2 = workbook.addImage({
      base64: myBase64Image,
      extension: 'png',
    })
    wsh.addImage(imageId2, 'J' + (8 + lengthSpisok) + ':K' + (11 + lengthSpisok));

    const buffer = await workbook.xlsx.writeBuffer();
    FileSaver.saveAs(new Blob([buffer]), `Заявка_УНИ_${dt}.xlsx`)
  }

  let [lengthSpisok, setLengthSpisok] = useState(1)

  //changeRulon()
  const addbut = document.getElementById('add')
  const xlsxbut = document.getElementById('xlsx')
  const tata = document.getElementById('ta')

  if (arrtabl.length > 0) {
    xlsxbut?.classList.remove('d-none')
  } else {
    xlsxbut?.classList.add('d-none')
  }

  if (w > 350 && w < 3300 && h > 350 && h < 3300 && num > 0 && num < 100) {
    addbut?.classList.remove('d-none')
  } else {
    addbut?.classList.add('d-none')
  }

  let [strokaToStr, setStrokaToStr] = useState('')

  const add = () => {
    setStrokaToStr(strokaToStr + '\n' + stroka.join(', ').replace(/УНИ/g, lengthSpisok + '. УНИ'))
    arrtabl.push(stroka)
    setArrtabl(arrtabl)
    //console.log(arrtabl)
    setLengthSpisok(arrtabl.length + 1)
    //let n = lengthSpisok + '. УНИ'
    //console.log(n)
    //setSpisok(arrtabl.join('\n').replace(/'УНИ'/g, n))
    //tata!.innerHTML = spisok
  }



  useEffect(() => {
    //setSpisok(arrtabl.join('\n').replace(/УНИ/g, lengthSpisok + '. УНИ'))
    setSpisok(strokaToStr)
  })

  return (


    <>
      <form className="row" noValidate>

        <div className="row mx-auto w-75 p-0">

          <div className=" col form-floating  form-control-sm">
            <input className="form-control" {...register("width", { required: true, min: 350, max: 3300, maxLength: 4 })} />
            <label>ширина, мм</label>
          </div>
          {errors.width && <span className="badge text-danger">введите ширину от 350мм до 3300мм</span>}

          <div className="col form-floating  form-control-sm">
            <input className="form-control" {...register("height", { required: true, min: 350, max: 3300, maxLength: 4 })} />
            <label>высота, мм</label>
          </div>
          {errors.height && <span className="badge text-danger">введите высоту от 350мм до 3300мм</span>}
        </div>


        <div className="container-fluid w-75 form-floating  form-control-sm">
          <select id='mat' className="form-select" defaultValue='АВЕНСИС' {...register("material")} onClick={onchangeFilter} onChange={onchangeFilter}>
            {SelectRulon}
          </select>
          <label>материал</label>
        </div>
        {errors.material && <span>материал</span>}

        <div className="container-fluid w-75 form-floating  form-control-sm" id="torender">
          <select id='col' className="form-select" defaultValue='белый' {...register("color")} onClick={onClo} onChange={onClo}>
            {cl}
          </select>
          <label>цвет</label>
        </div>
        {errors.color && <span>цвет</span>}

        <div className="row mx-auto w-75 p-0">

          <div className=" col form-floating  form-control-sm">
            <input className="form-control" defaultValue={1} {...register("num", { required: true, min: 1, max: 99, maxLength: 2 })} />
            <label>кол-во, шт</label>
          </div>
          {errors.num && <span>введите количество</span>}

          <div className="col form-floating  form-control-sm">
            <select id='upr' className="form-select" defaultValue="прав" {...register("upr")} onChange={onchanupr} onClick={onchanupr}>
              <option value="прав">прав</option>
              <option value="лев">лев</option>
              <option value="лев">лев/прав</option>
            </select>
            <label>управление</label>
          </div>
          {errors.upr && <span>управление</span>}

        </div>

        {/* <div className="d-flex justify-content-center mt-2">
          <input type="file" onChange={showFile}></input>
        </div> */}
        <div className="container-fluid w-100 form-floating  form-control-sm">
          <textarea className="form-control form-control-sm h-auto py-0 px-2" value={spisok} readOnly rows={lengthSpisok * 2}></textarea>
        </div>
      </form>



      <div className="row mx-auto w-75 p-0">

        <div className="col d-flex justify-content-center">
          <button id="add" className="btn btn-outline-primary mt-1 d-none" onClick={add}>++</button>
        </div>

        <div className="col d-flex justify-content-center">
          <button id="xlsx" className="btn btn-outline-primary mt-1 d-none" onClick={go}>xlsx</button>
        </div>

      </div>
    </>

  )

}