import 'regenerator-runtime/runtime'
let convertDocX = require('./docx2html')
let input = document.getElementById('input')
let box = document.getElementById('box')
input.onchange = async function (evt) {
  let file = input.files[0]
  box.innerHTML = await new convertDocX().buffer2html(await file2ArrayBuffer(file));
}
async function file2ArrayBuffer(file) {
  let buf
  await new Promise(resolve => {
    let fileReader = new FileReader()
    fileReader.onload = function () {
      buf = fileReader.result
      resolve()
    }
    fileReader.readAsArrayBuffer(file)
  })
  return buf
}