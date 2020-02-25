const docx = require("docx");

const { Document, Packer, Paragraph, Header, TextRun, HeadingLevel, VerticalAlign, Media } = docx;
const doc = new Document();
const fs = require("fs");

const image1 = Media.addImage(doc, fs.readFileSync("./image/image1.jpg"))

var docdata = {
    budgetyear: null,
    budgetsummary: 20000000,
    budgetinyear: 20000000,
    compcode: "01035",
    deptcode: "01035",
    plancode: "49",
    projectcode: "49003",
    activitycode: "49003001",
    sourcecode: "0000003",

    compname: "",
    deptname: "ศูนย์อำนวยการน้ำแห่งชาติ",
    groupname: "วิเคราะห์และติดตาม",
    planname: "แผนงานบูรณาการบริหารจัดการทรัพยากรน้ำ",
    projectname: "ขับเคลื่อนการดำเนินการตามกฎหมายและแผนแม่บทด้านการบริหารจัดการน้ำ",
    activityname: "จัดทำแผนและผลัดดันกฎหมายด้านการบริหาร",
    sourcename: "งบรายจ่ายอื่น",
    budgetsummarytext: "(ยี่สิบล้านบาทถ้วน)",

    name: "การประชุมเตรียมการสัปดาห์น้ำแห่งเอเชีย ครั้งที่ 2\tThe 2nd Asia International Water Week Preparatory  \n",
    owner: "\t\t\t\t\t\t\t\t\t  นางทยิดา สิริธีรธำรง ฟัน ก็อรสตันเยอ",
    criteria: "<p>สภาน้ำแห่งเอเชีย Asia Water Council  (AWC) มีวัตถุประสงค์เพื่อกระตุ้นและรองรับการเจริญเติบโตและการพัฒนาอย่างยั่งยืนของการบริหารจัดการน้ำในภูมิภาคเอเชีย โดยเน้นการมีส่วนร่วมของกลุ่มผู้มีส่วนได้ส่วนเสียอย่างทั่วถึง การสร้างความเข้าใจร่วมกันในประเด็นเกี่ยวกับการบริหารจัดการน้ำและการแก้ไขปัญหา และเพื่อก่อตั้งภาคีเครือข่ายความร่วมมือระหว่างประเทศสมาชิกในเอเชียและองค์กรนานาชาติอื่นๆ ซึ่งเลขาธิการสำนักงานทรัพยากรน้ำแห่งชาติ ได้ดำรงตำแหน่งเป็นสมาชิกกิตติมศักดิ์ (Honorary Membership) ของคณะกรรมการบริหารสภาน้ำแห่งเอเชีย เมื่อวันที่ 26 พฤษภาคม 2561 </p><p>การประชุม Asia International Water Week  จัดให้เป็นเวทีขององค์กรหรือกลุ่มต่างๆ ที่เกี่ยวกับด้านน้ำได้มีโอกาสมาพบปะแลกเปลี่ยนความรู้ ความก้าวหน้าในระดับประเทศ ระดับทวีป และระดับโลก นำไปสู่กรอบเป้าวัตถุประสงค์ร่วมกัน จุดมุ่งหมายคือผลักดันการพิจารณาปรับปรุงด้านน้ำให้เป็นวาระนโยบายในระดับต่างๆ ของแต่ละประเทศทั่วโลก ดังนั้น จึงควรให้มีการเข้าร่วมประชุมเพื่อรับทราบนโยบายและทิศทางในด้านน้ำ และเป็นการสนับสนุนกิจกรรมของ AWC</p><p></p><p></p>",
    objectives: "<p>&nbsp;&nbsp;&nbsp;&nbsp;1. เพื่อเตรียมการเข้าร่วมประชุม Asia International Water Week ครั้งที่ 2 เพื่อให้ทราบนโยบายและทิศทางของการจัดการน้ำในอนาคต</p><p>&nbsp;&nbsp;&nbsp;&nbsp;  2. เป็นการสนับสนุนกิจกรรมของ Asia Water Council (AWC)</p><p></p><p></p>",
    location: "<p>&nbsp;&nbsp;&nbsp;&nbsp;สาธารณรัฐอินโดนีเซีย</p><p></p><p></p>",
    targetgroup: "<p>ผู้บริหารและบุคลากรของสำนักงานทรัพยากรน้ำแห่งชาติและเจ้าหน้าที่หน่วยงานอื่นๆ ที่เกี่ยวข้อง จำนวน 2 คน</p><p></p><p></p>",
    timeline: "<p>ระยะเวลา 2 วัน ช่วงเดือนกรกฎาคม 2563</p><p></p><p></p>",
    process: "<p>กิจกรรม&nbsp;&nbsp;&nbsp;&nbsp;ไตรมาสที่ 1</p><p>(ต.ค. 62 - ธ.ค. 62)&nbsp;&nbsp;&nbsp;&nbsp;ไตรมาสที่ 2</p><p>(ม.ค. 63 - มี.ค. 63)&nbsp;&nbsp;&nbsp;&nbsp;ไตรมาสที่ 3</p><p>(เม.ย. 63 - มิ.ย. 63)&nbsp;&nbsp;&nbsp;&nbsp;ไตรมาสที่ 4</p><p>(ก.ค. 63 - ก.ย. 63)&nbsp;&nbsp;&nbsp;&nbsp;\tการประชุมภายใต้โครงการความร่วมมือ \t\t\t\t\t\t</p><p></p>",
    resulthistory: "<p>- </p><p></p>",
    budgetpaln: "วงเงินงบประมาณที่ขอในปี 2563 \t\n\n",
    output: "<p>ร้อยละของความสำเร็จของการดำเนินการตามแผน</p><p>ร้อยละของความสำเร็จของการเบิกจ่ายงบประมาณ</p><p>เอกสารทางวิชาการ รายงานการประชุม</p><p></p><p>14. ผลลัพธ์ของแผนงาน/โครงการ (Outcome)</p><p>1. ได้รับทราบทิศทางและแนวโน้มการจัดการน้ำของทั่วโลกในอนาคต</p><p>2. เพื่อให้สามารถดำเนินกิจกรรมของ AWC ได้อย่างต่อเนื่อง      </p><p></p><p>15. ผลประโยชน์ที่คาดว่าจะได้รับ</p><p>หน่วยงานภายในของ สทนช. ที่เกี่ยวข้องกับการกำหนดนโยบายและทิศทางการบริหารจัดการน้ำของประเทศ</p><p>หน่วยงานด้านน้ำมากกว่า 30 หน่วยงานทั่วประเทศ</p><p></p><p></p>",
    outcome: "<p>&nbsp;&nbsp;&nbsp;&nbsp;ผลการดำเนินงานสามารถเป็นกรอบแนวทางการปฏิบัติที่ดี สำหรับการขยายผลในหน่วยงานอื่นๆ ที่เกี่ยวข้องกับการบริหารจัดการทรัพยากรน้ำ </p><p></p><p></p><p></p>",
    benefit: "<p>แนวทางการเสนอขอโครงกา</p><p></p>",
    indicator: "",
    activities: [],
    budgetcodetmp: "01035490030000003XXXX",
    __v: 0
}

doc.addSection({

    // properties: {},
    headers: {
        default: new Header({
            children: [
                new Paragraph(image1),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "\t\t\t\t\t\tแบบประมาณการ",
                            bold: true,
                        }),
                    ],
                }),
            ],
        }),
    },
    children: [
        new Paragraph({
            children: [
                new TextRun({
                    text: "หน่วยงาน\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.deptname,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "กลุ่ม/ฝ่าย\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.groupname,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "รายการ\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.name,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "แผนงาน\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.planname,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ผลผลิต/โครงการ โครงการที่ 1\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.projectname,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "กิจกรรม กิจกรรมที่ 1.1\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.activityname,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ประเภทรายจ่าย\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.sourcename,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "รวมเงิน\t",
                    bold: true,
                }),
                new TextRun({
                    text: docdata.budgetsummary,
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun(docdata.budgetsummarytext),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun("\t\t\t--------------------------------------------------------------"),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "\t\t\t\t\t\tคำชี้แจง",
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun("เรียน ลทช. ผ่าน ผอ.ศอน."),
            ],
        }),
        new Paragraph({
            children: [
                // new TextRun("\tประมาณการฉบับนี้ตั้งขึ้นเพื่อควบคุมค่าใช้จ่ายในโครงการขับเคลื่อนนโยบาย และแผนแม่บทด้านการบริหารจัดการน้ำ"),
                new TextRun({
                    text: "\tประมาณการฉบับนี้ตั้งขึ้นเพื่อควบคุมค่าใช้จ่ายใน โครงการ",
                }),
                new TextRun({
                    text: docdata.name,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun("\tตามพระราชบัญญัติทรัพยากรน้ำ พ.ศ. 2561 กำหนดให้สำนักงานมีการจัดทำ 'ผังน้ำ' เพื่อเสนอให้กับคณะกรรมการทรัพยากรน้ำแห่งชาติ ภายในวันดังกล่าว"),
            ],
        }),
        new Paragraph({
            children: [
                // new TextRun("\tดังนั้น เพื่อให้การดำเนินงานบรรลุตามวัตุประสงค์ที่วางไว้ จึงขอ ทั้งสิ้น 55,278,300 บาท ( ห้าสิบห้าล้านสองแสนเจ็ดหมื่นแปดพันสามร้อยบาทถ้วน ) ตามรายละเอียด"),
                new TextRun({
                    text: "\tดังนั้น เพื่อให้การดำเนินงานบรรลุตามวัตุประสงค์ที่วางไว้ จึงขอ ทั้งสิ้น ",
                }),
                new TextRun({
                    text: docdata.budgetsummary,
                }),
                new TextRun({
                    text: docdata.budgetsummarytext,
                }),
                new TextRun({
                    text: "\tตามรายละเอียด",
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun("\t\t\tจึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ"),
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun(docdata.owner),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun("\t\t\t\t\t\t\t\t\tผู้อำนวยการกลุ่มวิเคราะห์และติดตามสถาน"),
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "\t\t\t  ผ่าน",
                    bold: true,
                }),
                new TextRun({
                    text: "\t\t\t\t\t   อนุมัติ",
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "\t\t      (นาย อุทัย เตียนพลกรัง)",
                    bold: true,
                }),
                new TextRun({
                    text: "\t\t\t      (นาย สมเกียรติ ประจำวงษ์)",
                    bold: true,
                }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "\t\tผู้อำนวยการศูนย์อำนวยการน้ำแห่งชาติ",
                    bold: true,
                }),
                new TextRun({
                    text: "\t\t\tเลขาธิการสำนักงานทรัพยากรน้ำแห่งชาติ",
                    bold: true,
                }),
            ],
        }),
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});