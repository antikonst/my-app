import React, { useState, useCallback } from "react";
import { useForm, SubmitHandler } from "react-hook-form";
import '../node_modules/bootstrap/dist/css/bootstrap.min.css'
import Excel from 'exceljs'

type Inputs = {
  width: number,
  height: number
};

export default function App() {
  const { register, handleSubmit, watch, formState: { errors } } = useForm<Inputs>();
  //const onSubmit: SubmitHandler<Inputs> = data => console.log(data);

  let w = watch("width")
  let h = watch("height")

  const [file, setFile] = useState('../src/blankUNI.xlsx')

const showFile = (e: React.FormEvent<HTMLInputElement>): void => {
  let files: FileList | null =  e.currentTarget.files;
  setFile(files![0].name)
  console.log(file)
}

const run = useCallback(async () => {
    const workbook = new Excel.Workbook();
    workbook.xlsx.readFile(file)
    let worksheet = workbook.getWorksheet('UNI')
    worksheet.getRow(-1).getCell('D').value = 350
    workbook.xlsx.writeFile('blankUNI.xlsx')
  }, [file])


  return (
    <form className="row" noValidate>

      <div className="container-fluid w-75 form-floating  form-control-sm">
        <input className="form-control" {...register("width", { required: true })} />
      </div>
      {errors.width && <span>введите ширину</span>}

      <div className="container-fluid w-75 form-floating  form-control-sm">
        <input className="form-control" {...register("height", { required: true })} />
      </div>
      {errors.height && <span>введите высоту</span>}

      <input type="file" onChange={showFile}></input>

      <div className="container-fluid w-75 form-floating  form-control-sm">
        <input className="btn btn-outline-primary position-relative mb-3" type="button" onClick={run}/>
      </div>
    </form>
  );
}