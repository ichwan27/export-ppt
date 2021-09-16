import { Button }  from 'antd';
import pptxgen from 'pptxgenjs';
import './App.css';

function ExportPPT(){
  let pptx =  new pptxgen();
  let slide1  = pptx.addSlide();
  let slide2 = pptx.addSlide();
    
    // judul 1
     slide1.addText(
      "Judul Gambar 1",
      {
      x: 0.3,
      y: 0.1,
      w: 2,
      h: 0.1,
      fontSize: 14,
    });
    //border kiri 1
    slide1.addShape(pptx.shapes.RECTANGLE, {
      x: 0.1,
      y: 0.29,
      w: 2.04,
      h: 1.11,
      line: { color: "000000", width: 1 },
    });
      // kiri 1
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 0.13,
        y: 0.32,
        w: 2,
        h: 1.06,
      });
    // judul 2
    slide1.addText(
      "Judul Gambar 2",
      {
      x: 0.3,
      y: 1.5,
      w: 2,
      h: 0.1,
      fontSize: 14,
    });
      //border kiri 2
      slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 0.08,
        y: 1.67,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // gambar 2 kiri 2
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 0.1,
        y: 1.70,
        w: 2,
        h: 1.06,
      });
       // judul 3
       slide1.addText(
        "Judul Gambar 3",
       {
        x: 0.3,
        y: 2.9,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       
      //border kiri 3
      slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 0.08,
        y: 3.07,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // kiri 3
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 0.1,
        y: 3.10,
        w: 2,
        h: 1.06,
      });
      // judul 4
      slide1.addText(
      "Judul Gambar 4",
      {
      x: 0.3,
      y: 4.3,
      w: 2,
      h: 0.1,
      fontSize: 14,
      });
       //border kiri 4
       slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 0.08,
        y: 4.47,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // kiri 4
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 0.1,
        y: 4.50,
        w: 2,
        h: 1.06,
      });
      // judul tengah 1
     slide1.addText(
      "Judul Gambar 5",
      {
      x: 2.4,
      y: 0.1,
      w: 2,
      h: 0.1,
      fontSize: 14,
     });
      //border tengah 1
      slide1.addShape(pptx.shapes.RECTANGLE, {
      x: 2.20,
      y: 0.29,
      w: 2.04,
      h: 1.11,
      line: { color: "000000", width: 1 },
    });
      // tengah 1
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 2.23,
        y: 0.32,
        w: 2,
        h: 1.06,
        });
      // judul tengah 2
      slide1.addText(
        "Judul Gambar 6",
        {
        x: 2.4,
        y: 1.5,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border tengah 2
        slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 2.20,
        y: 1.67,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // tengah 2
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 2.23,
        y: 1.70,
        w: 2,
        h: 1.06,
      });
       // judul tengah 3
       slide1.addText(
        "Judul Gambar 7",
        {
        x: 2.4,
        y: 2.9,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border tengah 3
       slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 2.20,
        y: 3.07,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // tengah 3
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 2.23,
        y: 3.10,
        w: 2,
        h: 1.06,
      });
      // judul tengah 4
      slide1.addText(
        "Judul Gambar 8",
        {
        x: 2.4,
        y: 4.3,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
      // border tengah 4
      slide1.addShape(pptx.shapes.RECTANGLE, {
          x: 2.20,
          y: 4.47,
          w: 2.04,
          h: 1.11,
          line: { color: "000000", width: 1 },
        });
      // tengah 4
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 2.23,
        y: 4.5,
        w: 2,
        h: 1.06,
      });
      // judul kanan 1
      slide1.addText(
        "Judul Gambar 9",
        {
        x: 4.5,
        y: 0.1,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border kanan 1
       slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 4.3,
        y: 0.30,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
       //kanan 1
       slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 4.33,
        y: 0.32,
        w: 2,
        h: 1.06,
      });
      // judul kanan 2
      slide1.addText(
        "Judul Gambar 10",
        {
        x: 4.5,
        y: 1.5,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
        //border kanan 2
        slide1.addShape(pptx.shapes.RECTANGLE, {
          x: 4.3,
          y: 1.68,
          w: 2.04,
          h: 1.11,
          line: { color: "000000", width: 1 },
        });
      // kanan 2
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 4.33,
        y: 1.7,
        w: 2,
        h: 1.06,
      });
      // judul kanan 3
      slide1.addText(
        "Judul Gambar 11",
        {
        x: 4.5,
        y: 2.9,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border kanan 3
       slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 4.3,
        y: 3.08,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // kanan 3
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 4.33,
        y: 3.1,
        w: 2,
        h: 1.06,
      });
      // judul kanan 4
      slide1.addText(
        "Judul Gambar 9",
        {
        x: 4.5,
        y: 4.3,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
      //border kanan 4
      slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 4.3,
        y: 4.48,
        w: 2.04,
        h: 1.11,
        line: { color: "000000", width: 1 },
      });
      // kanan 4
      slide1.addImage({
        path: "https://quickchart.io/chart?bkg=white&c={type:%27pie%27,data:{labels:[%27January%27,%27February%27,%27March%27,%27April%27,%20%27May%27],datasets:[{data:[50,60,70,180,190]}]}}",
        x: 4.33,
        y: 4.5,
        w: 2,
        h: 1.06,
      });
      //judul gambar kanan bawah
      slide1.addText(
        "Judul Gambar 10",
        {
        x: 7.4,
        y: 3.8,
        w: 2,
        h: 0.1,
        fontSize: 14,
        bold: true,
       });
       //border kanan bawah
      slide1.addShape(pptx.shapes.RECTANGLE, {
        x: 6.42,
        y: 3.95,
        w: 3.55,
        h: 1.56,
        line: { color: "000000", width: 1 },
      });
      //Kanan Bawah
      slide1.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT8pw68IoiAi4tIFz5taJlgSWf3LiWgPmKowg&usqp=CAU",
        x: 6.45,
        y: 4,
        w: 3.5,
        h: 1.5,
      });
      // komen
      let objOptions = {
        x: 6.4,
        y: 0.1,
        w: 3.58,
        h: 0.3,
        bold: true,
        align: "center",
        fontSize: 16,
        color: "ffffff",
        fill: { color: "24207a" },
      };
      let objOptions2 = {
        x: 6.4,
        y: 0.4,
        w: 3.58,
        h: 0.3,
        bold: true,
        align: "center",
        fontSize: 16,
        margin: 10,
        fill: { color: "F1F1F1" },
      };
      slide1.addText("PERFORMANCE", objOptions);
      slide1.addText("All Area Comment", objOptions2);
      slide1.addTable(
        [
          [
            {
              text: "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.",
            },
          ],
        ],
        {
          x: 6.4,
          y: 0.7,
          w: 3.58,
          h: 3,
          fontSize: 12,
          border: { color: "9c9c9c" },
        },
      );
    //Judul slide 2
    slide2.addText(
      "Resources Report",
      {
      x: 0.2,
      y: 0.25,
      w: 9.45,
      h: 0.64,
      fontSize: 28,
      bold: true,
      align: "center",
    });
    // 2 judul 2
    slide2.addText(
      "Judul Gambar 2",
      {
      x: 1.3,
      y: 1,
      w: 2,
      h: 0.1,
      fontSize: 14,
    });
      //border 2 kiri 1
      slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 1,
        y: 1.17,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // gambar 2 kiri 1
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
        x: 1.03,
        y: 1.20,
        w: 2,
        h: 1.2,
      });
       // 2 judul 2
       slide2.addText(
        "Judul Gambar 3",
       {
        x: 1.3,
        y: 2.47,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       
      //border 2 kiri 2
      slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 1,
        y: 2.64,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // 2 kiri 2
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
        x: 1.03,
        y: 2.67,
        w: 2,
        h: 1.2,
      });
      // 2 judul 3
      slide2.addText(
      "Judul Gambar 4",
      {
      x: 1.3,
      y: 4.04,
      w: 2,
      h: 0.1,
      fontSize: 14,
      });
       //border 2 kiri 3
       slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 1,
        y: 4.21,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // 2 kiri 3
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
        x: 1.03,
        y: 4.24,
        w: 2,
        h: 1.2,
      });
      // judul 2 tengah 1
      slide2.addText(
        "Judul Gambar 6",
        {
        x: 4.3,
        y: 1,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border 2 tengah 1
        slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 4,
        y: 1.17,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // 2 tengah 1
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
        x: 4.03,
        y: 1.20,
        w: 2,
        h: 1.2,
      });
       // judul 2 tengah 2
       slide2.addText(
        "Judul Gambar 7",
        {
        x: 4.3,
        y: 2.47,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border 2 tengah 2
       slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 4,
        y: 2.64,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // 2 tengah 2
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
        x: 4.03,
        y: 2.67,
        w: 2,
        h: 1.2,
      });
      // judul 2 tengah 3
      slide2.addText(
        "Judul Gambar 8",
        {
        x: 4.3,
        y: 4.04,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
      // border 2 tengah 3
      slide2.addShape(pptx.shapes.RECTANGLE, {
         x: 4,
         y: 4.21,
         w: 2.04,
         h: 1.24,
          line: { color: "000000", width: 1 },
        });
      // 2 tengah 3
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
        x: 4.03,
        y: 4.24,
        w: 2,
        h: 1.2,
      });
      // judul 2 kanan 1
      slide2.addText(
        "Judul Gambar 10",
        {
        x: 7.3,
        y: 1,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
        //border 2 kanan 1
        slide2.addShape(pptx.shapes.RECTANGLE, {
          x: 7,
          y: 1.17,
          w: 2.04,
          h: 1.24,
          line: { color: "000000", width: 1 },
        });
      //  2 kanan 1
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
        x: 7.03,
        y: 1.20,
        w: 2,
        h: 1.2,
      });
      // judul kanan 2
      slide2.addText(
        "Judul Gambar 11",
        {
        x: 7.3,
        y: 2.47,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
       //border 2 kanan 2
       slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 7,
         y: 2.64,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // 2 kanan 2
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
        x: 7.03,
        y: 2.67,
        w: 2,
        h: 1.2,
      });
      // judul 2 kanan 3
      slide2.addText(
        "Judul Gambar 9",
        {
        x: 7.3,
        y: 4.04,
        w: 2,
        h: 0.1,
        fontSize: 14,
       });
      //border 2 kanan 3
      slide2.addShape(pptx.shapes.RECTANGLE, {
        x: 7,
        y: 4.21,
        w: 2.04,
        h: 1.24,
        line: { color: "000000", width: 1 },
      });
      // 2 kanan 3
      slide2.addImage({
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",    
        x: 7.03,
        y: 4.24,
        w: 2,
        h: 1.2,
      });
        
  pptx.writeFile({ fileName: 'NSEM-Report-Analytic-.pptx' });
}

function App() {
  return (
    <div
    className="background"
    >
      <div
      className="top-field" 
      >
        <h2>NSEM Report Analytic Dashboard</h2>
          <Button
          onClick={ExportPPT}
          className="btn-export"
          >
          Export to PPT
          </Button>
      </div>
      <div className="report-div">
        <h3>Performance Report</h3>
      </div>
      <div className="report-div">
        <h3>Resources Report</h3>
      </div>
    </div>
  );
}

export default App;
