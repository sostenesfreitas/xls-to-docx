import React from 'react';
import './App.css';
import readXlsxFile from 'read-excel-file';
import { saveAs } from 'file-saver';
import { Spinner } from 'react-bootstrap';
import { Document, Paragraph, Packer, TextRun } from 'docx';
function App() {
  let fileInput = React.createRef();
  const [spinner, setSpinner] = React.useState(false);
  function documentGenerate(xls) {
    readXlsxFile(xls.current.files[0], { getSheets: true })
      .then(rows => {
        let prom = rows.map(sheet => {
          return readXlsxFile(xls.current.files[0], { sheet: sheet.name });
        });
        Promise.all(prom).then(data => {
          let test = [];
          data.forEach(d => {
            d.forEach(e => {
              test.push(e);
            });
          });
          test = test.map(t => {
            return createParagraph(t);
          });
          generateDoc(test);
        });
      })
      .catch(error => {
        console.log(error);
      });
  }

  const createParagraph = row => {
    const paragraph1 = new Paragraph({
      children: [
        new TextRun({
          text: 'Número do Caso de Teste:',
          bold: true
        }).break(),
        new TextRun({
          text: row[1]
        }),
        new TextRun({
          text: 'Prova:',
          bold: true
        }).break(),
        new TextRun({
          text:
            '[Incluir a tela, a tabela ou arquivo contendo o resultado obtido e comentários que facilitem a avaliação do conteúdo da prova]',
          color: '0000FF',
          size: 20,
          italics: true
        }),
        new TextRun({
          text: 'Resultado de Teste:',
          bold: true
        }).break(),
        new TextRun({
          text: row[5]
        }).break()
      ]
    });
    return paragraph1;
  };

  const generateDoc = paragraphs => {
    const doc = new Document();
    doc.addSection({
      children: paragraphs
    });
    Packer.toBlob(doc).then(blob => {
      saveAs(blob, 'example.docx');
      console.log('Document created successfully');
      setSpinner(false);
    });
  };
  function treta(f) {
    setSpinner(true);
  }
  return (
    <div className='App'>
      <div>
        <input
          type='file'
          id='input'
          ref={fileInput}
          onClick={() => setSpinner(true)}
          onChange={() => documentGenerate(fileInput)}
        />
        <div className='spinner'>
          {!spinner ? (
            <div></div>
          ) : (
            <>
              <Spinner
                as='span'
                animation='grow'
                size='sm'
                role='status'
                aria-hidden='true'
              />
              Loading...
            </>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
