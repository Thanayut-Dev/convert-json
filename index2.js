const docx = require("docx");
const { Document, Packer, Paragraph, Header, TextRun, AlignmentType, VerticalAlign } = docx;
const fs = require("fs");

const doc = new Document({

    styles: {
        paragraphStyles: [
            {
                id: "heading-end",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    size: 28,
                    bold: true,
                },
                paragraph: {
                    alignment: AlignmentType.END,
                },
            },
            {
                id: "heading-center",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    size: 38,
                    bold: true,
                },
                paragraph: {
                    alignment: AlignmentType.CENTER,
                },
            },
            {
                id: "format-title",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    size: 28,
                    bold: true,
                },
                paragraph: {
                    alignment: AlignmentType.START,
                },
            },
            {
                id: "format-subtitle",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    size: 28,
                },
                paragraph: {
                    alignment: AlignmentType.START,
                },
            },
        ]
    }
});

var docxData001 = {
    budgetyear: null,
    budgetsummary: 20000000,
    budgetinyear: 20000000,
    compcode: "01035",
    deptcode: "01035",
    plancode: "49",
    projectcode: "49003",
    activitycode: "49003001",
    sourcecode: "0000003",
    name: "การประชุมเตรียมการสัปดาห์น้ำแห่งเอเชีย ครั้งที่ 2 The 2nd Asia International Water Week Preparatory",
    owner: "\tชื่อ-นามสกุล \tนางทยิดา สิริธีรธำรง ฟัน ก็อรสตันเยอ \n\tตำแหน่ง \tนักวิเทศสัมพันธ์ชำนาญการพิเศษ สังกัด \tกองการต่างประเทศ โทรศัพท์เคลื่อนที่ \t06 5519 6055 E-mail address \tthayida@gmail.com",
    criteria: "\tสภาน้ำแห่งเอเชีย Asia Water Council (AWC) มีวัตถุประสงค์เพื่อกระตุ้นและรองรับการเจริญเติบโตและการพัฒนาอย่างยั่งยืนของการบริหารจัดการน้ำในภูมิภาคเอเชีย โดยเน้นการมีส่วนร่วมของกลุ่มผู้มีส่วนได้ส่วนเสียอย่างทั่วถึง การสร้างความเข้าใจร่วมกันในประเด็นเกี่ยวกับการบริหารจัดการน้ำและการแก้ไขปัญหา และเพื่อก่อตั้งภาคีเครือข่ายความร่วมมือระหว่างประเทศสมาชิกในเอเชียและองค์กรนานาชาติอื่นๆ ซึ่งเลขาธิการสำนักงานทรัพยากรน้ำแห่งชาติ ได้ดำรงตำแหน่งเป็นสมาชิกกิตติมศักดิ์ (Honorary Membership) ของคณะกรรมการบริหารสภาน้ำแห่งเอเชีย เมื่อวันที่ 26 พฤษภาคม 2561 การประชุม Asia International Water Week  จัดให้เป็นเวทีขององค์กรหรือกลุ่มต่างๆ ที่เกี่ยวกับด้านน้ำได้มีโอกาสมาพบปะแลกเปลี่ยนความรู้ ความก้าวหน้าในระดับประเทศ ระดับทวีป และระดับโลก นำไปสู่กรอบเป้าวัตถุประสงค์ร่วมกัน จุดมุ่งหมายคือผลักดันการพิจารณาปรับปรุงด้านน้ำให้เป็นวาระนโยบายในระดับต่างๆ ของแต่ละประเทศทั่วโลก ดังนั้น จึงควรให้มีการเข้าร่วมประชุมเพื่อรับทราบนโยบายและทิศทางในด้านน้ำ และเป็นการสนับสนุนกิจกรรมของ AWC",
    objectives: "\t1. เพื่อเตรียมการเข้าร่วมประชุม Asia International Water Week ครั้งที่ 2 เพื่อให้ทราบนโยบายและทิศทางของการจัดการน้ำในอนาคต\t2. เป็นการสนับสนุนกิจกรรมของ Asia Water Council (AWC)",
    location: "\tสาธารณรัฐอินโดนีเซีย",
    targetgroup: "\tผู้บริหารและบุคลากรของสำนักงานทรัพยากรน้ำแห่งชาติและเจ้าหน้าที่หน่วยงานอื่นๆ ที่เกี่ยวข้อง จำนวน 2 คน",
    timeline: "\tระยะเวลา 2 วัน ช่วงเดือนกรกฎาคม 2563",
    process: "\tกิจกรรม ไตรมาสที่ 1 (ต.ค. 62 - ธ.ค. 62) ไตรมาสที่ 2 (ม.ค. 63 - มี.ค. 63) ไตรมาสที่ 3 (เม.ย. 63 - มิ.ย. 63) ไตรมาสที่ 4 (ก.ค. 63 - ก.ย. 63) \tการประชุมภายใต้โครงการความร่วมมือ \t\t\t\t\t\t ",
    resulthistory: "\t-",
    budgetpaln: "\tวงเงินงบประมาณที่ขอในปี 2563 \t\n\n",
    output: "\tร้อยละของความสำเร็จของการดำเนินการตามแผน ร้อยละของความสำเร็จของการเบิกจ่ายงบประมาณ เอกสารทางวิชาการ รายงานการประชุม 14. ผลลัพธ์ของแผนงาน/โครงการ (Outcome) 1. ได้รับทราบทิศทางและแนวโน้มการจัดการน้ำของทั่วโลกในอนาคต 2. เพื่อให้สามารถดำเนินกิจกรรมของ AWC ได้อย่างต่อเนื่อง 15. ผลประโยชน์ที่คาดว่าจะได้รับ หน่วยงานภายในของ สทนช. ที่เกี่ยวข้องกับการกำหนดนโยบายและทิศทางการบริหารจัดการน้ำของประเทศ หน่วยงานด้านน้ำมากกว่า 30 หน่วยงานทั่วประเทศ",
    outcome: "\tผลการดำเนินงานสามารถเป็นกรอบแนวทางการปฏิบัติที่ดี สำหรับการขยายผลในหน่วยงานอื่นๆ ที่เกี่ยวข้องกับการบริหารจัดการทรัพยากรน้ำ",
    benefit: "\tแนวทางการเสนอขอโครงกา",
    indicator: "\tใช้ผังน้ำประกอบการวางแผนบริหารจัดการทรัพยากรน้ำอย่างเป็นระบบ",
    activities: [],
    budgetcodetmp: "01035490030000003XXXX",
    __v: 0
}

doc.addSection({

    headers: {
        default: new Header({
            children: [
                new Paragraph({
                    text: "แบบฟอร์ม กผง.001",
                    style: "heading-end",
                }),
            ],
        }),
    },

    children: [
        new Paragraph({
            text: "ข้อเสนอโครงการที่จะเสนอขอตั้งงบประมาณรายจ่ายประจำปีงบประมาณ พ.ศ. 2563 สำนักงานทรัพยากรน้ำแห่งชาติ",
            style: "heading-center",
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
                    text: "ชื่อโครงการ :\t",
                    bold: true,
                }),
                new TextRun({
                    text: docxData001.name,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ผู้รับผิดชอบ :\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.owner,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "หลักการและเหตุผล\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.criteria,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "วัตถุประสงค์\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.objectives,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "พื้นที่ดำเนินโครงการ\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.location,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "กลุ่มเป้าหมาย ผู้มีส่วนได้ส่วนเสีย\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.targetgroup,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ระยะเวลาดำเนินโครงการ\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.timeline,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "แผนการปฏิบัติงาน/วิธีการดำเนินงาน/กิจกรรม (โดยละเอียด)\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.process,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ผลการดำเนินงานที่ผ่านมา (ถ้ามี)\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.resulthistory,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "งบประมาณ และแผนการใช้จ่ายงบประมาณ\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.budgetpaln,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ผลผลิตของแผนงาน/โครงการ (Output) และตัวชี้วัดของโครงการ\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.output,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ผลลัพธ์/ผลสัมฤทธิ์ของแผนงาน/โครงการ (Outcome)\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.outcome,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ผลประโยชน์/ผลกระทบที่คาดว่าจะได้รับ\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.benefit,
                }),
            ],
            style: "format-subtitle",
        }),
        new Paragraph({
            children: [
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "การติดตามและประเมินผลโครงการ\t",
                    bold: true,
                }),
            ],
            style: "format-title",
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: docxData001.indicator,
                }),
            ],
            style: "format-subtitle",
        }),
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.doc", buffer);
});