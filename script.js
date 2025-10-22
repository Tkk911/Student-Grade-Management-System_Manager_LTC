// ฐาน URL ของ Google Apps Script Web App
const GAS_URL = 'https://script.google.com/macros/s/AKfycbydEBkMT09_QEm6yM43oGyjzzuHYingSpXvyFrwRT-LpZCp_a-GNqpGqyww24UxuFxyDw/exec';

// ข้อมูลหลักสูตรและสาขาจาก Excel - อัพเดทให้ตรงกับข้อมูลจริง
const programData = {
    'ปวช.': {
        'ช่างไฟฟ้า': {
            semesters: 6,
            courses: {
                '1': [
                    { code: "20000-1101", name: "ภาษาไทยเพื่อการสื่อสาร", credits: 1 },
                    { code: "20000-1201", name: "ภาษาอังกฤษเพื่อการสื่อสาร", credits: 1 },
                    { code: "20000-1301", name: "วิทยาศาสตร์พื้นฐานอาชีพ", credits: 2 },
                    { code: "20000-1401", name: "คณิตศาสตร์พื้นฐานอาชีพ", credits: 2 },
                    { code: "20000-1502", name: "ประวัติศาสตร์ชาติไทย", credits: 1 },
                    { code: "20000-1603", name: "พลศึกษาเพื่อพัฒนาสุขภาพ", credits: 1 },
                    { code: "20001-1005", name: "การใช้เทคโนโลยีดิจิทัลเพื่ออาชีพ", credits: 3 },
                    { code: "20100-1001", name: "เขียนแบบเทคนิคเบื้องต้น", credits: 2 },
                    { code: "20100-1003", name: "งานฝึกฝีมือ", credits: 2 },
                    { code: "20100-1005", name: "งานไฟฟ้าและอิเล็กทรอนิกส์เบื้องต้น", credits: 2 },
                    { code: "20104-2017", name: "กฎและมาตรฐานทางไฟฟ้า", credits: 2 },
                    { code: "20000-2001", name: "กิจกรรมลูกเสือวิสามัญ 1", credits: 0 }
                ],
                '2': [
                    { code: "20000-1102", name: "ภาษาไทยเพื่ออาชีพ", credits: 1 },
                    { code: "20000-1203", name: "การฟังและการพูดภาษาอังกฤษ", credits: 1 },
                    { code: "20000-1302", name: "วิทยาศาสตร์เพื่ออาชีพอุตสาหกรรม", credits: 2 },
                    { code: "20001-1001", name: "สุขภาพความปลอดภัยและสิ่งแวดล้อม", credits: 2 },
                    { code: "20100-1004", name: "งานเชื่อมและโลหะแผ่นเบื้องต้น", credits: 2 },
                    { code: "20100-1006", name: "งานเครื่องมือกลเบื้องต้น", credits: 2 },
                    { code: "20104-2001", name: "เขียนแบบไฟฟ้า", credits: 2 },
                    { code: "20104-2004", name: "เครื่องวัดไฟฟ้า", credits: 2 },
                    { code: "20104-2005", name: "การติดตั้งไฟฟ้าในอาคาร", credits: 3 },
                    { code: "20000-2002", name: "กิจกรรมลูกเสือวิสามัญ 2", credits: 0 }
                ],
                '3': [
                    { code: "20000-1104", name: "การใช้ภาษาไทยในยุคดิจิทัล", credits: 1 },
                    { code: "20000-1204", name: "ภาษาอังกฤษสถานประกอบการ", credits: 1 },
                    { code: "20000-1501", name: "หน้าที่พลเมืองและศีลธรรม", credits: 2 },
                    { code: "20000-1602", name: "เพศวิถีศึกษา", credits: 1 },
                    { code: "20104-2002", name: "วงจรไฟฟ้ากระแสตรง", credits: 2 },
                    { code: "20104-2006", name: "เครื่องกลไฟฟ้ากระแสตรง", credits: 2 },
                    { code: "20104-2007", name: "เครื่องทำความเย็น", credits: 3 },
                    { code: "20104-2018", name: "อุปกรณ์อิเล็กทรอนิกส์และวงจร", credits: 2 },
                    { code: "20102-2002", name: "เขียนแบบด้วยโปรแกรมคอมพิวเตอร์", credits: 2 },
                    { code: "20104-2025", name: "เทคนิคการจัดการพลังงาน", credits: 2 },
                    { code: "20000-2003", name: "กิจกรรมเสริมสร้างสุจริต จิตอาสา", credits: 0 }
                ],
                '4': [
                    { code: "20000-1105", name: "การใช้ภาษาไทยเชิงสร้างสรรค์", credits: 1 },
                    { code: "20000-1209", name: "ภาษาอังกฤษเพื่องานช่างไฟฟ้าและอิเล็กทรอนิกส์", credits: 1 },
                    { code: "20001-1002", name: "การพัฒนาอย่างยั่งยืน", credits: 2 },
                    { code: "20100-1007", name: "งานนิวเมติกส์และไฮดรอลิกส์เบื้องต้น", credits: 2 },
                    { code: "20104-2003", name: "วงจรไฟฟ้ากระแสสลับ", credits: 2 },
                    { code: "20104-2008", name: "มอเตอร์ไฟฟ้ากระแสสลับ", credits: 3 },
                    { code: "20104-2013", name: "หม้อแปลงไฟฟ้า", credits: 2 },
                    { code: "20104-2015", name: "เครื่องปรับอากาศ", credits: 3 },
                    { code: "20000-2004", name: "กิจกรรมองค์การวิชาชีพ 1", credits: 0 }
                ],
                '5': [
                    { code: "20104-2009", name: "การควบคุมมอเตอร์ไฟฟ้า", credits: 3 },
                    { code: "20104-2010", name: "การประมาณการติดตั้งไฟฟ้า", credits: 2 },
                    { code: "20104-2011", name: "การโปรแกรมและควบคุมไฟฟ้า", credits: 2 },
                    { code: "20104-2014", name: "เครื่องกำเนิดไฟฟ้ากระแสสลับ", credits: 2 },
                    { code: "20104-2024", name: "เครื่องวัดอุตสาหกรรมและควบคุมเบื้องต้น", credits: 2 },
                    { code: "20000-2007", name: "กิจกรรมในสถานประกอบการ 1", credits: 0 }
                ],
                '6': [
                    { code: "20000-1221", name: "ภาษาอังกฤษเพื่อเตรียมความพร้อมเพื่อการทำงาน", credits: 1 },
                    { code: "20001-1003", name: "ธุรกิจเบื้องต้น", credits: 2 },
                    { code: "20001-1004", name: "กฎหมายแรงงาน", credits: 1 },
                    { code: "20104-2012", name: "การติดตั้งไฟฟ้านอกอาคาร", credits: 3 },
                    { code: "20104-2021", name: "ไมโครคอนโทรลเลอร์เบื้องต้น", credits: 2 },
                    { code: "20104-2029", name: "โครงงานด้านช่างไฟฟ้า", credits: 4 },
                    { code: "20104-2023", name: "การส่องสว่าง", credits: 2 },
                    { code: "20104-2027", name: "คณิตศาสตร์ไฟฟ้า", credits: 2 },
                    { code: "20104-2016", name: "งานซ่อมเครื่องใช้ไฟฟ้า", credits: 2 },
                    { code: "20000-2005", name: "กิจกรรมองค์การวิชาชีพ 2", credits: 0 }
                ]
            }
        },
        'อิเล็กทรอนิกส์': {
            semesters: 6,
            courses: {
                '1': [
                    { code: "20000-1101", name: "ภาษาไทยเพื่อการสื่อสาร", credits: 1 },
                    { code: "20000-1201", name: "ภาษาอังกฤษเพื่อการสื่อสาร", credits: 1 },
                    { code: "20000-1301", name: "วิทยาศาสตร์พื้นฐานอาชีพ", credits: 2 },
                    { code: "20000-1401", name: "คณิตศาสตร์พื้นฐานอาชีพ", credits: 2 },
                    { code: "20000-1502", name: "ประวัติศาสตร์ชาติไทย", credits: 1 },
                    { code: "20000-1603", name: "พลศึกษาเพื่อพัฒนาสุขภาพ", credits: 1 },
                    { code: "20001-1005", name: "การใช้เทคโนโลยีดิจิทัลเพื่ออาชีพ", credits: 3 },
                    { code: "20100-1001", name: "เขียนแบบเทคนิคเบื้องต้น", credits: 2 },
                    { code: "20100-1003", name: "งานฝึกฝีมือ", credits: 2 },
                    { code: "20100-1005", name: "งานไฟฟ้าและอิเล็กทรอนิกส์เบื้องต้น", credits: 2 },
                    { code: "20000-2001", name: "กิจกรรมลูกเสือวิสามัญ 1", credits: 0 }
                ],
                '2': [
                    { code: "20000-1102", name: "ภาษาไทยเพื่ออาชีพ", credits: 1 },
                    { code: "20000-1203", name: "การฟังและการพูดภาษาอังกฤษ", credits: 1 },
                    { code: "20000-1302", name: "วิทยาศาสตร์เพื่ออาชีพอุตสาหกรรม", credits: 2 },
                    { code: "20001-1001", name: "สุขภาพความปลอดภัยและสิ่งแวดล้อม", credits: 2 },
                    { code: "20100-1004", name: "งานเชื่อมและโลหะแผ่นเบื้องต้น", credits: 2 },
                    { code: "20100-1006", name: "งานเครื่องมือกลเบื้องต้น", credits: 2 },
                    { code: "20105-2001", name: "วงจรไฟฟ้า", credits: 3 },
                    { code: "20105-2002", name: "อุปกรณ์อิเล็กทรอนิกส์และวงจร", credits: 3 },
                    { code: "20105-2023", name: "เครื่องมือวัดไฟฟ้าและอิเล็กทรอนิกส์", credits: 2 },
                    { code: "20000-2002", name: "กิจกรรมลูกเสือวิสามัญ 2", credits: 0 }
                ],
                '3': [
                    { code: "20000-1104", name: "การใช้ภาษาไทยในยุคดิจิทัล", credits: 1 },
                    { code: "20000-1204", name: "ภาษาอังกฤษสถานประกอบการ", credits: 1 },
                    { code: "20000-1501", name: "หน้าที่พลเมืองและศีลธรรม", credits: 2 },
                    { code: "20000-1602", name: "เพศวิถีศึกษา", credits: 1 },
                    { code: "20105-2003", name: "วงจรพัลส์และดิจิทัล", credits: 2 },
                    { code: "20105-2004", name: "เขียนแบบอิเล็กทรอนิกส์ด้วยคอมพิวเตอร์", credits: 3 },
                    { code: "20105-2008", name: "วงจรไอซีและการประยุกต์ใช้งาน", credits: 2 },
                    { code: "20105-2011", name: "เครื่องเสียง", credits: 2 },
                    { code: "20105-2017", name: "เครื่องส่งวิทยุ", credits: 2 },
                    { code: "20105-2016", name: "โทรทัศน์ระบบดิจิทัล", credits: 3 },
                    { code: "20000-2003", name: "กิจกรรมเสริมสร้างสุจริต จิตอาสา", credits: 0 }
                ],
                '4': [
                    { code: "20000-1105", name: "การใช้ภาษาไทยเชิงสร้างสรรค์", credits: 1 },
                    { code: "20000-1209", name: "ภาษาอังกฤษเพื่องานช่างไฟฟ้าและอิเล็กทรอนิกส์", credits: 1 },
                    { code: "20001-1002", name: "การพัฒนาอย่างยั่งยืน", credits: 2 },
                    { code: "20100-1007", name: "งานนิวเมติกส์และไฮดรอลิกส์เบื้องต้น", credits: 2 },
                    { code: "20105-2005", name: "ไมโครคอนโทรลเลอร์", credits: 3 },
                    { code: "20105-2007", name: "การเขียนโปรแกรมคอมพิวเตอร์", credits: 3 },
                    { code: "20105-2019", name: "เครื่องรับวิทยุ", credits: 2 },
                    { code: "20105-2021", name: "อินเตอร์เฟซเบื้องต้น", credits: 2 },
                    { code: "20105-2024", name: "คณิตศาสตร์ช่างอิเล็กทรอนิกส์", credits: 2 },
                    { code: "20000-2004", name: "กิจกรรมองค์การวิชาชีพ 1", credits: 0 }
                ],
                '5': [
                    { code: "20105-2009", name: "อิเล็กทรอนิกส์อุตสาหกรรม", credits: 2 },
                    { code: "20105-2014", name: "เครือข่ายคอมพิวเตอร์", credits: 2 },
                    { code: "20105-2018", name: "สายส่งและสายอากาศ", credits: 2 },
                    { code: "20105-2020", name: "งานบริการอิเล็กทรอนิกส์", credits: 2 },
                    { code: "20105-2025", name: "อุปกรณ์อิเล็กทรอนิกส์ในระบบรักษาความปลอดภัย", credits: 2 },
                    { code: "20000-2007", name: "กิจกรรมในสถานประกอบการ 1", credits: 0 }
                ],
                '6': [
                    { code: "20000-1221", name: "ภาษาอังกฤษเพื่อเตรียมความพร้อมเพื่อการทำงาน", credits: 1 },
                    { code: "20001-1003", name: "ธุรกิจเบื้องต้น", credits: 2 },
                    { code: "20001-1004", name: "กฎหมายแรงงาน", credits: 1 },
                    { code: "20105-2006", name: "โปรแกรมเมเบิลลอจิกคอนโทรล", credits: 3 },
                    { code: "20105-2010", name: "หุ่นยนต์เบื้องต้น", credits: 2 },
                    { code: "20105-2032", name: "โครงงานด้านอิเล็กทรอนิกส์", credits: 4 },
                    { code: "20105-2022", name: "พื้นฐานเซนเซอร์ในงานอุตสาหกรรม", credits: 3 },
                    { code: "20105-2026", name: "งานบริการคอมพิวเตอร์", credits: 2 },
                    { code: "20000-2005", name: "กิจกรรมองค์การวิชาชีพ 2", credits: 0 }
                ]
            }
        }
    },
    'ปวส.': {
        'ไฟฟ้า': {
            'ปวส.ชฟ.ห้อง1': {
                semesters: 4,
                courses: {
                    '1': [
                        { code: "30000-1101", name: "ทักษะภาษาไทยเพื่อการสื่อสารในงานอาชีพ", credits: 2 },
                        { code: "30000-1201", name: "ภาษาอังกฤษสำหรับงานอาชีพ", credits: 2 },
                        { code: "30000-1404", name: "แคลคูลัส 1", credits: 3 },
                        { code: "30000-1503", name: "หลักปรัชญาของเศรษฐกิจพอเพียงเพื่อการดำเนินชีวิต", credits: 1 },
                        { code: "30001-1001", name: "การเป็นผู้ประกอบการ", credits: 3 },
                        { code: "30104-2001", name: "เครื่องมือวัดไฟฟ้า", credits: 3 },
                        { code: "30104-2002", name: "วงจรไฟฟ้า", credits: 3 },
                        { code: "30104-2011", name: "การประมาณการระบบไฟฟ้า", credits: 3 },
                        { code: "30104-2066", name: "เทคนิคการจัดการความปลอดภัยในงานไฟฟ้า", credits: 2 },
                        { code: "30000-2001", name: "กิจกรรมเสริมสร้างสุจริต จิตอาสา", credits: 0 }
                    ],
                    '2': [
                        { code: "30000-1102", name: "ทักษะการเขียนและการพูดภาษาไทยในงานอาชีพ", credits: 2 },
                        { code: "30000-1302", name: "วิทยาศาสตร์งานอาชีพไฟฟ้า อิเล็กทรอนิกส์ และการสื่อสาร", credits: 3 },
                        { code: "30000-1606", name: "ภาวะผู้นำและการทำงานเป็นทีม", credits: 2 },
                        { code: "30001-0003", name: "การประยุกต์ใช้เทคโนโลยีดิจิทัลในอาชีพ", credits: 3 },
                        { code: "30100-1003", name: "กฎหมายในงานอาชีพอุตสาหกรรมพลังงาน ไฟฟ้า และอิเล็กทรอนิกส์", credits: 1 },
                        { code: "30104-2003", name: "การติดตั้งไฟฟ้า 1", credits: 3 },
                        { code: "30104-2004", name: "การออกแบบระบบไฟฟ้า", credits: 3 },
                        { code: "30104-2005", name: "เครื่องกลไฟฟ้า 1", credits: 3 },
                        { code: "30104-2046", name: "ระบบไฟฟ้าสำรอง", credits: 2 },
                        { code: "30000-2002", name: "กิจกรรมองค์การวิชาชีพ 1", credits: 0 }
                    ],
                    '3': [
                        { code: "30104-2007", name: "ระบบควบคุมในงานอุตสาหกรรม", credits: 3 },
                        { code: "30104-2020", name: "เครื่องกลไฟฟ้า 2", credits: 3 },
                        { code: "30104-2028", name: "ระบบปรับอากาศในงานอุตสาหกรรม", credits: 3 },
                        { code: "30104-2063", name: "เซลล์แสงอาทิตย์และการประยุกต์ใช้", credits: 3 },
                        { code: "30104-2026", name: "การติดตั้งไฟฟ้า 2", credits: 3 },
                        { code: "30000-2005", name: "กิจกรรมในสถานประกอบการ 1", credits: 0 }
                    ],
                    '4': [
                        { code: "30001-1002", name: "องค์การและการบริหารงานคุณภาพ", credits: 3 },
                        { code: "30100-1020", name: "การควบคุมนิวส์แมติกส์และไฮดรอลิกส์", credits: 3 },
                        { code: "30104-2008", name: "การเขียนโปรแกรมคอมพิวเตอร์ในงานควบคุมไฟฟ้า", credits: 3 },
                        { code: "30104-2006", name: "การเขียนแบบไฟฟ้าด้วยคอมพิวเตอร์", credits: 3 },
                        { code: "30104-2070", name: "โครงงานด้านไฟฟ้า", credits: 4 },
                        { code: "30104-2034", name: "การส่งและจ่ายไฟฟ้า", credits: 3 },
                        { code: "30104-2009", name: "คณิตศาสตร์ไฟฟ้า", credits: 3 },
                        { code: "30000-2003", name: "กิจกรรมองค์การวิชาชีพ 2", credits: 0 }
                    ]
                }
            },
            'ปวส.ชฟ.ห้อง6': {
                semesters: 4,
                courses: {
                    '1': [
                        { code: "30104-0001", name: "เทคนิคและเครื่องมือกลพื้นฐานสำหรับงานไฟฟ้า", credits: 3 },
                        { code: "30104-0003", name: "เครื่องกลไฟฟ้าและการควบคุม", credits: 3 },
                        { code: "30104-2066", name: "เทคนิคการจัดการความปลอดภัยในงานไฟฟ้า", credits: 2 },
                        { code: "30000-1101", name: "ทักษะภาษาไทยเพื่อการสื่อสารในงานอาชีพ", credits: 2 },
                        { code: "30000-1201", name: "ภาษาอังกฤษสำหรับงานอาชีพ", credits: 2 },
                        { code: "30000-1404", name: "แคลคูลัส 1", credits: 3 },
                        { code: "30000-1503", name: "หลักปรัชญาของเศรษฐกิจพอเพียงเพื่อการดำเนินชีวิต", credits: 1 },
                        { code: "30001-1001", name: "การเป็นผู้ประกอบการ", credits: 3 },
                        { code: "30104-2001", name: "เครื่องมือวัดไฟฟ้า", credits: 3 },
                        { code: "30104-2002", name: "วงจรไฟฟ้า", credits: 3 },
                        { code: "30104-2011", name: "การประมาณการระบบไฟฟ้า", credits: 3 },
                        { code: "30000-2001", name: "กิจกรรมเสริมสร้างสุจริต จิตอาสา", credits: 0 }
                    ],
                    '2': [
                        { code: "30104-0002", name: "การเขียนแบบไฟฟ้าและประมาณราคา", credits: 2 },
                        { code: "30104-0004", name: "การติดตั้งไฟฟ้าในและนอกอาคาร", credits: 3 },
                        { code: "30104-0005", name: "เครื่องทำความเย็นและปรับอากาศ", credits: 3 },
                        { code: "30000-1102", name: "ทักษะการเขียนและการพูดภาษาไทยในงานอาชีพ", credits: 2 },
                        { code: "30000-1302", name: "วิทยาศาสตร์งานอาชีพไฟฟ้า อิเล็กทรอนิกส์ และการสื่อสาร", credits: 3 },
                        { code: "30000-1606", name: "ภาวะผู้นำและการทำงานเป็นทีม", credits: 2 },
                        { code: "30001-0003", name: "การประยุกต์ใช้เทคโนโลยีดิจิทัลในอาชีพ", credits: 3 },
                        { code: "30100-1003", name: "กฎหมายในงานอาชีพอุตสาหกรรมพลังงาน ไฟฟ้า และอิเล็กทรอนิกส์", credits: 1 },
                        { code: "30104-2003", name: "การติดตั้งไฟฟ้า 1", credits: 3 },
                        { code: "30104-2004", name: "การออกแบบระบบไฟฟ้า", credits: 3 },
                        { code: "30104-2005", name: "เครื่องกลไฟฟ้า 1", credits: 3 },
                        { code: "30104-2046", name: "ระบบไฟฟ้าสำรอง", credits: 2 },
                        { code: "30000-2002", name: "กิจกรรมองค์การวิชาชีพ 1", credits: 0 }
                    ],
                    '3': [
                        { code: "30104-2007", name: "ระบบควบคุมในงานอุตสาหกรรม", credits: 3 },
                        { code: "30104-2020", name: "เครื่องกลไฟฟ้า 2", credits: 3 },
                        { code: "30104-2026", name: "การติดตั้งไฟฟ้า 2", credits: 3 },
                        { code: "30104-2028", name: "ระบบปรับอากาศในงานอุตสาหกรรม", credits: 3 },
                        { code: "30104-2063", name: "เซลล์แสงอาทิตย์และการประยุกต์ใช้", credits: 3 },
                        { code: "30000-2005", name: "กิจกรรมในสถานประกอบการ 1", credits: 0 }
                    ],
                    '4': [
                        { code: "30001-1002", name: "องค์การและการบริหารงานคุณภาพ", credits: 3 },
                        { code: "30100-1020", name: "การควบคุมนิวส์แมติกส์และไฮดรอลิกส์", credits: 3 },
                        { code: "30104-2006", name: "การเขียนแบบไฟฟ้าด้วยคอมพิวเตอร์", credits: 3 },
                        { code: "30104-2008", name: "การเขียนโปรแกรมคอมพิวเตอร์ในงานควบคุมไฟฟ้า", credits: 3 },
                        { code: "30104-2009", name: "คณิตศาสตร์ไฟฟ้า", credits: 3 },
                        { code: "30104-2034", name: "การส่งและจ่ายไฟฟ้า", credits: 3 },
                        { code: "30104-2070", name: "โครงงานด้านไฟฟ้า", credits: 4 },
                        { code: "30000-2003", name: "กิจกรรมองค์การวิชาชีพ 2", credits: 0 }
                    ]
                }
            }
        },
        'เทคโนโลยีอิเล็กทรอนิกส์': {
            'ปวส.ชอ.ห้อง1': {
                semesters: 4,
                courses: {
                    '1': [
                        { code: "30000-1101", name: "ทักษะภาษาไทยเพื่อการสื่อสารในงานอาชีพ", credits: 2 },
                        { code: "30000-1201", name: "ภาษาอังกฤษสำหรับงานอาชีพ", credits: 2 },
                        { code: "30000-1404", name: "แคลคูลัส 1", credits: 3 },
                        { code: "30000-1503", name: "หลักปรัชญาของเศรษฐกิจพอเพียงเพื่อการดำเนินชีวิต", credits: 1 },
                        { code: "30001-1001", name: "การเป็นผู้ประกอบการ", credits: 3 },
                        { code: "30105-2001", name: "การออกแบบวงจรอิเล็กทรอนิกส์ด้วยคอมพิวเตอร์", credits: 3 },
                        { code: "30105-2029", name: "วงจรพัลส์และดิจิทัลเทคนิค", credits: 3 },
                        { code: "30105-2030", name: "การวิเคราะห์วงจรอิเล็กทรอนิกส์", credits: 3 },
                        { code: "30105-2022", name: "เครื่องมือและอุปกรณ์ทางการแพทย์พื้นฐาน", credits: 2 },
                        { code: "30000-2001", name: "กิจกรรมเสริมสร้างสุจริต จิตอาสา", credits: 0 }
                    ],
                    '2': [
                        { code: "30000-1102", name: "ทักษะการเขียนและการพูดภาษาไทยในงานอาชีพ", credits: 2 },
                        { code: "30000-1302", name: "วิทยาศาสตร์งานอาชีพไฟฟ้า อิเล็กทรอนิกส์ และการสื่อสาร", credits: 3 },
                        { code: "30000-1606", name: "ภาวะผู้นำและการทำงานเป็นทีม", credits: 2 },
                        { code: "30001-0003", name: "การประยุกต์ใช้เทคโนโลยีดิจิทัลในอาชีพ", credits: 3 },
                        { code: "30105-2002", name: "สมองกลฝังตัว", credits: 3 },
                        { code: "30105-2004", name: "เทคนิคการอินเตอร์เฟซ", credits: 3 },
                        { code: "30105-2006", name: "การเขียนโปรแกรมคอมพิวเตอร์", credits: 3 },
                        { code: "30105-2027", name: "ความปลอดภัยอาชีวอนามัยและสิ่งแวดล้อมในการปฏิบัติงาน", credits: 3 },
                        { code: "30000-2002", name: "กิจกรรมองค์การวิชาชีพ 1", credits: 0 }
                    ],
                    '3': [
                        { code: "30105-2005", name: "อิเล็กทรอนิกส์อุตสาหกรรม", credits: 3 },
                        { code: "30105-2007", name: "หุ่นยนต์อุตสาหกรรม", credits: 3 },
                        { code: "30105-2016", name: "ระบบสื่อสารด้วยเส้นใยแก้วนำแสง", credits: 2 },
                        { code: "30105-2018", name: "ระบบเครือข่ายคอมพิวเตอร์", credits: 3 },
                        { code: "30105-2031", name: "เทคโนโลยีพลังงานทดแทนอัจฉริยะ", credits: 3 },
                        { code: "30000-2005", name: "กิจกรรมในสถานประกอบการ 1", credits: 0 }
                    ],
                    '4': [
                        { code: "30001-1002", name: "องค์การและการบริหารงานคุณภาพ", credits: 3 },
                        { code: "30100-1020", name: "การควบคุมนิวส์แมติกส์และไฮดรอลิกส์", credits: 3 },
                        { code: "30100-1003", name: "กฎหมายในงานอาชีพอุตสาหกรรมพลังงาน ไฟฟ้า และอิเล็กทรอนิกส์", credits: 1 },
                        { code: "30105-2003", name: "โปรแกรมเมเบิลลอจิกคอนโทรล", credits: 3 },
                        { code: "30105-2011", name: "โปรแกรมจำลองระบบการผลิตอัตโนมัติ", credits: 3 },
                        { code: "30105-2008", name: "การพัฒนาระบบควบคุมอัตโนมัติในงานอุตสาหกรรม", credits: 3 },
                        { code: "30105-2012", name: "อินเทอร์เน็ตของสรรพสิ่ง", credits: 2 },
                        { code: "30105-2032", name: "โครงงานด้านเทคโนโลยีอิเล็กทรอนิกส์", credits: 4 },
                        { code: "30000-2003", name: "กิจกรรมองค์การวิชาชีพ 2", credits: 0 }
                    ]
                }
            },
            'ปวส.ชอ.ห้อง6': {
                semesters: 4,
                courses: {
                    '1': [
                        { code: "30100-0010", name: "เทคนิคและเครื่องมือกลพื้นฐาน", credits: 3 },
                        { code: "30105-0001", name: "เขียนแบบอิเล็กทรอนิกส์ด้วยคอมพิวเตอร์", credits: 3 },
                        { code: "30000-1101", name: "ทักษะภาษาไทยเพื่อการสื่อสารในงานอาชีพ", credits: 2 },
                        { code: "30000-1201", name: "ภาษาอังกฤษสำหรับงานอาชีพ", credits: 2 },
                        { code: "30000-1404", name: "แคลคูลัส 1", credits: 3 },
                        { code: "30000-1503", name: "หลักปรัชญาของเศรษฐกิจพอเพียงเพื่อการดำเนินชีวิต", credits: 1 },
                        { code: "30001-1001", name: "การเป็นผู้ประกอบการ", credits: 3 },
                        { code: "30105-2001", name: "การออกแบบวงจรอิเล็กทรอนิกส์ด้วยคอมพิวเตอร์", credits: 3 },
                        { code: "30105-2029", name: "วงจรพัลส์และดิจิทัลเทคนิค", credits: 3 },
                        { code: "30105-2030", name: "การวิเคราะห์วงจรอิเล็กทรอนิกส์", credits: 3 },
                        { code: "30105-2022", name: "เครื่องมือและอุปกรณ์ทางการแพทย์พื้นฐาน", credits: 2 },
                        { code: "30000-2001", name: "กิจกรรมเสริมสร้างสุจริต จิตอาสา", credits: 0 }
                    ],
                    '2': [
                        { code: "30105-0002", name: "งานพื้นฐานวงจรไฟฟ้าและการวัด", credits: 3 },
                        { code: "30105-0003", name: "งานพื้นฐานวงจรอิเล็กทรอนิกส์", credits: 3 },
                        { code: "30105-0004", name: "งานพื้นฐานไฟฟ้าและอิเล็กทรอนิกส์", credits: 3 },
                        { code: "30000-1102", name: "ทักษะการเขียนและการพูดภาษาไทยในงานอาชีพ", credits: 2 },
                        { code: "30000-1302", name: "วิทยาศาสตร์งานอาชีพไฟฟ้า อิเล็กทรอนิกส์ และการสื่อสาร", credits: 3 },
                        { code: "30000-1606", name: "ภาวะผู้นำและการทำงานเป็นทีม", credits: 2 },
                        { code: "30001-0003", name: "การประยุกต์ใช้เทคโนโลยีดิจิทัลในอาชีพ", credits: 3 },
                        { code: "30105-2002", name: "สมองกลฝังตัว", credits: 3 },
                        { code: "30105-2004", name: "เทคนิคการอินเตอร์เฟซ", credits: 3 },
                        { code: "30105-2006", name: "การเขียนโปรแกรมคอมพิวเตอร์", credits: 3 },
                        { code: "30105-2027", name: "ความปลอดภัยอาชีวอนามัยและสิ่งแวดล้อมในการปฏิบัติงาน", credits: 3 },
                        { code: "30000-2002", name: "กิจกรรมองค์การวิชาชีพ 1", credits: 0 }
                    ],
                    '3': [
                        { code: "30105-2005", name: "อิเล็กทรอนิกส์อุตสาหกรรม", credits: 3 },
                        { code: "30105-2007", name: "หุ่นยนต์อุตสาหกรรม", credits: 3 },
                        { code: "30105-2016", name: "ระบบสื่อสารด้วยเส้นใยแก้วนำแสง", credits: 2 },
                        { code: "30105-2018", name: "ระบบเครือข่ายคอมพิวเตอร์", credits: 3 },
                        { code: "30105-2031", name: "เทคโนโลยีพลังงานทดแทนอัจฉริยะ", credits: 3 },
                        { code: "30000-2005", name: "กิจกรรมในสถานประกอบการ 1", credits: 0 }
                    ],
                    '4': [
                        { code: "30001-1002", name: "องค์การและการบริหารงานคุณภาพ", credits: 3 },
                        { code: "30100-1020", name: "การควบคุมนิวส์แมติกส์และไฮดรอลิกส์", credits: 3 },
                        { code: "30100-1003", name: "กฎหมายในงานอาชีพอุตสาหกรรมพลังงาน ไฟฟ้า และอิเล็กทรอนิกส์", credits: 1 },
                        { code: "30105-2003", name: "โปรแกรมเมเบิลลอจิกคอนโทรล", credits: 3 },
                        { code: "30105-2011", name: "โปรแกรมจำลองระบบการผลิตอัตโนมัติ", credits: 3 },
                        { code: "30105-2008", name: "การพัฒนาระบบควบคุมอัตโนมัติในงานอุตสาหกรรม", credits: 3 },
                        { code: "30105-2012", name: "อินเทอร์เน็ตของสรรพสิ่ง", credits: 2 },
                        { code: "30105-2032", name: "โครงงานด้านเทคโนโลยีอิเล็กทรอนิกส์", credits: 4 },
                        { code: "30000-2003", name: "กิจกรรมองค์การวิชาชีพ 2", credits: 0 }
                    ]
                }
            }
        }
    }
};

// DOM Elements
const guestMode = document.getElementById('guestMode');
const adminMode = document.getElementById('adminMode');
const gradeSheet = document.getElementById('gradeSheet');
const adminGradeSheet = document.getElementById('adminGradeSheet');
const modeIndicator = document.getElementById('modeIndicator');
const modeText = document.getElementById('modeText');
const loginBtn = document.getElementById('loginBtn');
const loginText = document.getElementById('loginText');
const printBtn = document.getElementById('printBtn');
const searchBtn = document.getElementById('searchBtn');
const adminSearchBtn = document.getElementById('adminSearchBtn');
const addStudentBtn = document.getElementById('addStudentBtn');
const saveAllBtn = document.getElementById('saveAllBtn');
const importBtn = document.getElementById('importBtn');
const exportBtn = document.getElementById('exportBtn');
const loginModal = document.getElementById('loginModal');
const studentModal = document.getElementById('studentModal');
const courseModal = document.getElementById('courseModal');
const importExportModal = document.getElementById('importExportModal');
const loginForm = document.getElementById('loginForm');
const studentForm = document.getElementById('studentForm');
const courseForm = document.getElementById('courseForm');
const searchInput = document.getElementById('searchInput');
const searchStudent = document.getElementById('searchStudent');

// ตัวแปรสถานะ
let currentMode = 'guest'; // 'guest' หรือ 'admin'
let currentStudent = null;
let nextCourseId = 1000; // สำหรับเพิ่มวิชาใหม่
let nextStudentId = 27000; // สำหรับเพิ่มนักเรียนใหม่

// Event Listeners
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupEventListeners();
});

function initializeApp() {
    // ตั้งค่าโหมดเริ่มต้นเป็นผู้เยี่ยมชม
    setMode('guest');
    // ตั้งค่าธีม
    initializeTheme();
}

function setupEventListeners() {
    // Print button
    printBtn.addEventListener('click', function() {
        window.print();
    });
    
    // Login button
    loginBtn.addEventListener('click', function() {
        if (currentMode === 'guest') {
            showLoginModal();
        } else {
            logout();
        }
    });
    
    // Search buttons
    searchBtn.addEventListener('click', function() {
        searchStudentInGuestMode();
    });
    
    adminSearchBtn.addEventListener('click', function() {
        searchStudentInAdminMode();
    });
    
    // Add student button
    addStudentBtn.addEventListener('click', function() {
        showStudentModal();
    });
    
    // Import/Export buttons
    importBtn.addEventListener('click', function() {
        showImportExportModal('import');
    });
    
    exportBtn.addEventListener('click', function() {
        showImportExportModal('export');
    });
    
    // Save all button
    saveAllBtn.addEventListener('click', function() {
        saveAllGrades();
    });
    
    // Modal close buttons
    document.querySelectorAll('.close-modal').forEach(button => {
        button.addEventListener('click', function() {
            this.closest('.modal').style.display = 'none';
        });
    });
    
    // Close modal when clicking outside
    document.querySelectorAll('.modal').forEach(modal => {
        modal.addEventListener('click', function(e) {
            if (e.target === this) {
                this.style.display = 'none';
            }
        });
    });
    
    // Forms
    loginForm.addEventListener('submit', handleLogin);
    studentForm.addEventListener('submit', handleStudentSubmit);
    courseForm.addEventListener('submit', handleCourseSubmit);
    
    // Import/Export tab buttons
    document.querySelectorAll('.modal-tab').forEach(tab => {
        tab.addEventListener('click', function() {
            const tabName = this.dataset.tab;
            switchImportExportTab(tabName);
        });
    });
    
    // Import/Export action buttons
    document.getElementById('importDataBtn').addEventListener('click', importData);
    document.getElementById('exportDataBtn').addEventListener('click', exportData);
    
    // Enter key in search inputs
    searchInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            searchStudentInGuestMode();
        }
    });
    
    searchStudent.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            searchStudentInAdminMode();
        }
    });
    
    // Filter changes
    document.getElementById('filterClass').addEventListener('change', applyFilters);
    document.getElementById('filterSemester').addEventListener('change', applyFilters);
    document.getElementById('filterProgram').addEventListener('change', applyFilters);
    document.getElementById('filterMajor').addEventListener('change', applyFilters);
    
    // Add courses from program button
    document.getElementById('addCoursesFromProgramBtn').addEventListener('click', showAddCoursesModal);
    
    // Add courses modal buttons
    document.getElementById('cancelAddCourses').addEventListener('click', function() {
        document.getElementById('addCoursesModal').style.display = 'none';
    });
    
    document.getElementById('confirmAddCourses').addEventListener('click', confirmAddCourses);
    
    // Semester change in add courses modal
    document.getElementById('coursesSemester').addEventListener('change', function() {
        if (document.getElementById('addCoursesModal').style.display === 'flex') {
            showAddCoursesModal();
        }
    });
}

// ==================== THEME MANAGEMENT ====================

function initializeTheme() {
    const savedTheme = localStorage.getItem('theme') || 'light';
    setTheme(savedTheme);
    
    // Event listener สำหรับปุ่มสลับธีม
    document.getElementById('themeToggle').addEventListener('click', toggleTheme);
}

function setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
    
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) {
        if (theme === 'dark') {
            themeToggle.innerHTML = '<i class="fas fa-sun"></i><span>โหมดกลางวัน</span>';
        } else {
            themeToggle.innerHTML = '<i class="fas fa-moon"></i><span>โหมดกลางคืน</span>';
        }
    }
}

function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    setTheme(newTheme);
}

// ==================== MODE MANAGEMENT ====================

function setMode(mode) {
    currentMode = mode;
    
    if (mode === 'guest') {
        guestMode.style.display = 'block';
        adminMode.style.display = 'none';
        modeText.textContent = 'โหมดผู้เยี่ยมชม';
        loginText.textContent = 'เข้าสู่ระบบ';
        loginBtn.className = 'btn btn-primary';
        modeIndicator.className = 'mode-indicator';
    } else {
        guestMode.style.display = 'none';
        adminMode.style.display = 'block';
        modeText.textContent = 'โหมดผู้ดูแล';
        loginText.textContent = 'ออกจากระบบ';
        loginBtn.className = 'btn btn-danger';
        modeIndicator.className = 'mode-indicator admin';
    }
}

// ==================== LOGIN MANAGEMENT ====================

function showLoginModal() {
    loginModal.style.display = 'flex';
    document.getElementById('username').focus();
}

async function handleLogin(e) {
    e.preventDefault();
    
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    
    // แสดงสถานะกำลังโหลด
    const submitBtn = loginForm.querySelector('button[type="submit"]');
    const originalText = submitBtn.innerHTML;
    submitBtn.innerHTML = '<div class="loading"></div> กำลังเข้าสู่ระบบ...';
    submitBtn.disabled = true;
    
    const result = await callGAS('login', {
        username: username,
        password: password
    });
    
    // คืนค่าปุ่มเป็นปกติ
    submitBtn.innerHTML = originalText;
    submitBtn.disabled = false;
    
    if (result.success) {
        setMode('admin');
        loginModal.style.display = 'none';
        loginForm.reset();
        showNotification('เข้าสู่ระบบผู้ดูแลสำเร็จ', 'success');
    } else {
        showNotification('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง', 'error');
    }
}

function logout() {
    setMode('guest');
    showNotification('ออกจากระบบผู้ดูแลแล้ว', 'info');
}

// ==================== STUDENT SEARCH ====================

async function searchStudentInGuestMode() {
    const query = searchInput.value.trim();
    
    if (!query) {
        showNotification('กรุณากรอกรหัสหรือชื่อนักเรียน', 'warning');
        return;
    }
    
    const result = await callGAS('searchStudents', { query: query });
    
    if (result.success && result.data.length > 0) {
        const student = result.data[0];
        // ดึงข้อมูลรายวิชาของนักเรียน
        const studentDetail = await callGAS('getStudentById', { id: student.id });
        
        if (studentDetail.success) {
            loadStudentDataInGuestMode(studentDetail.data);
            showNotification(`พบข้อมูลนักเรียน: ${student.name}`, 'success');
        } else {
            showNotification('ไม่พบข้อมูลผลการเรียน', 'error');
        }
    } else {
        showNotification('ไม่พบข้อมูลนักเรียน', 'error');
        gradeSheet.style.display = 'none';
    }
}

async function searchStudentInAdminMode() {
    const query = searchStudent.value.trim();
    
    if (!query) {
        showNotification('กรุณากรอกรหัสหรือชื่อนักเรียน', 'warning');
        return;
    }
    
    const result = await callGAS('searchStudents', { query: query });
    
    if (result.success && result.data.length > 0) {
        const student = result.data[0];
        // ดึงข้อมูลรายวิชาของนักเรียน
        const studentDetail = await callGAS('getStudentById', { id: student.id });
        
        if (studentDetail.success) {
            currentStudent = studentDetail.data;
            loadStudentDataInAdminMode(currentStudent);
            showNotification(`พบข้อมูลนักเรียน: ${student.name}`, 'success');
        } else {
            showNotification('ไม่พบข้อมูลผลการเรียน', 'error');
        }
    } else {
        showNotification('ไม่พบข้อมูลนักเรียน', 'error');
    }
}

// ==================== GRADE DISPLAY ====================

function loadStudentDataInGuestMode(student) {
    currentStudent = student;
    
    // สร้าง HTML สำหรับแสดงผลการเรียน (โหมดอ่านอย่างเดียว)
    let html = `
        <div class="sheet-header">
            <h1>ใบแสดงผลการเรียน</h1>
            <p>วิทยาลัยเทคโนโลยีแหลมทอง • ประเภทวิชาอุตสาหกรรม • สาขาวิชา ${student.major} • หลักสูตร${student.program} ${student.year}</p>
        </div>
        
        <div class="student-info">
            <div class="student-details">
                <h3>รหัส: ${student.id} ชื่อ: ${student.name}</h3>
                <p>ระดับชั้น: ${student.level} | สาขา: ${student.major} | ห้อง: ${student.class}</p>
            </div>
            <div class="gpa-summary">
                <p>เกรดเฉลี่ยสะสม</p>
                <div class="gpa-value">${calculateGPA(student).toFixed(2)}</div>
            </div>
        </div>
    `;
    
    // เพิ่มภาคเรียนตามหลักสูตร
    const semestersCount = student.program === 'ปวช.' ? 6 : 4;
    
    // สร้างภาคเรียนเป็นคู่
    for (let i = 1; i <= semestersCount; i += 2) {
        html += createSemesterPairForGuest(i.toString(), (i + 1).toString(), student);
    }
    
    gradeSheet.innerHTML = html;
    gradeSheet.style.display = 'block';
}

function loadStudentDataInAdminMode(student) {
    currentStudent = student;
    
    let html = `
        <div class="sheet-header">
            <h1>ใบแสดงผลการเรียน</h1>
            <p>วิทยาลัยเทคโนโลยีแหลมทอง • ประเภทวิชาอุตสาหกรรม • สาขาวิชา ${student.major} • หลักสูตร${student.program} ${student.year}</p>
        </div>
        
        <div class="student-info">
            <div class="student-details">
                <h3>รหัส: ${student.id} ชื่อ: ${student.name}</h3>
                <p>ระดับชั้น: ${student.level} | สาขา: ${student.major} | ห้อง: ${student.class}</p>
            </div>
            <div class="gpa-summary">
                <p>เกรดเฉลี่ยสะสม</p>
                <div class="gpa-value">${calculateGPA(student).toFixed(2)}</div>
            </div>
        </div>
    `;
    
    // กำหนดจำนวนภาคเรียนตามหลักสูตร
    const semestersCount = student.program === 'ปวช.' ? 6 : 4;
    
    // สร้างภาคเรียนเป็นคู่
    for (let i = 1; i <= semestersCount; i += 2) {
        html += createSemesterPairForAdmin(i.toString(), (i + 1).toString(), student);
    }
    
    adminGradeSheet.innerHTML = html;
    
    // เพิ่ม event listener สำหรับ dropdown เกรด
    setTimeout(() => {
        document.querySelectorAll('.grade-select').forEach(select => {
            select.addEventListener('change', function() {
                const courseId = this.dataset.courseId;
                const semester = this.dataset.semester;
                const grade = this.value;
                
                updateGradeInDatabase(courseId, semester, grade);
            });
        });
    }, 100);
}

function createSemesterPairForGuest(semester1, semester2, student) {
    const courses1 = student.grades[semester1] || [];
    const courses2 = student.grades[semester2] || [];
    const credits1 = calculateCredits(courses1);
    const credits2 = calculateCredits(courses2);
    
    return `
        <div class="semester-pair">
            <div class="semester-pair-header">
                <span class="semester-pair-title">ภาคเรียนที่ ${semester1} และ ภาคเรียนที่ ${semester2}</span>
            </div>
            
            <div class="semester-pair-content">
                <!-- Semester ${semester1} -->
                <div class="semester-column">
                    <table class="grades-table">
                        <thead>
                            <tr>
                                <th colspan="4" style="text-align: center;">ภาคเรียนที่ ${semester1}</th>
                            </tr>
                            <tr>
                                <th width="120">รหัสวิชา</th>
                                <th>ชื่อวิชา</th>
                                <th width="60">นก.</th>
                                <th width="80">เกรด</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${courses1.map(course => `
                                <tr>
                                    <td>${course.code}</td>
                                    <td>${course.name}</td>
                                    <td>${course.credits}</td>
                                    <td class="grade-cell grade-${course.grade ? course.grade.replace('.', '\\.') : ''}">
                                        ${course.grade || '-'}
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    <div class="credits-summary">
                        หน่วยกิตรวม: ${credits1}
                    </div>
                </div>
                
                <!-- Semester ${semester2} -->
                <div class="semester-column">
                    <table class="grades-table">
                        <thead>
                            <tr>
                                <th colspan="4" style="text-align: center;">ภาคเรียนที่ ${semester2}</th>
                            </tr>
                            <tr>
                                <th width="120">รหัสวิชา</th>
                                <th>ชื่อวิชา</th>
                                <th width="60">นก.</th>
                                <th width="80">เกรด</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${courses2.map(course => `
                                <tr>
                                    <td>${course.code}</td>
                                    <td>${course.name}</td>
                                    <td>${course.credits}</td>
                                    <td class="grade-cell grade-${course.grade ? course.grade.replace('.', '\\.') : ''}">
                                        ${course.grade || '-'}
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    <div class="credits-summary">
                        หน่วยกิตรวม: ${credits2}
                    </div>
                </div>
            </div>
        </div>
    `;
}

function createSemesterPairForAdmin(semester1, semester2, student) {
    const courses1 = student.grades[semester1] || [];
    const courses2 = student.grades[semester2] || [];
    const credits1 = calculateCredits(courses1);
    const credits2 = calculateCredits(courses2);
    
    return `
        <div class="semester-pair">
            <div class="semester-pair-header">
                <span class="semester-pair-title">ภาคเรียนที่ ${semester1} และ ภาคเรียนที่ ${semester2}</span>
                <div class="semester-pair-actions">
                    <button class="btn btn-success btn-sm" onclick="addCourseToSemester('${semester1}')">
                        <i class="fas fa-plus"></i> เพิ่มวิชา
                    </button>
                </div>
            </div>
            
            <div class="semester-pair-content">
                <!-- Semester ${semester1} -->
                <div class="semester-column">
                    <table class="grades-table">
                        <thead>
                            <tr>
                                <th colspan="5" style="text-align: center;">ภาคเรียนที่ ${semester1}</th>
                            </tr>
                            <tr>
                                <th width="120">รหัสวิชา</th>
                                <th>ชื่อวิชา</th>
                                <th width="60">นก.</th>
                                <th width="80">เกรด</th>
                                <th width="100" class="actions-cell">การจัดการ</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${courses1.map(course => `
                                <tr>
                                    <td>${course.code}</td>
                                    <td>${course.name}</td>
                                    <td>${course.credits}</td>
                                    <td class="grade-cell">
                                        <select class="form-control grade-select" data-course-id="${course.id}" data-semester="${semester1}" style="padding: 4px; font-size: 0.8rem;">
                                            <option value="">-- เลือก --</option>
                                            <option value="4" ${course.grade === '4' ? 'selected' : ''}>4</option>
                                            <option value="3.5" ${course.grade === '3.5' ? 'selected' : ''}>3.5</option>
                                            <option value="3" ${course.grade === '3' ? 'selected' : ''}>3</option>
                                            <option value="2.5" ${course.grade === '2.5' ? 'selected' : ''}>2.5</option>
                                            <option value="2" ${course.grade === '2' ? 'selected' : ''}>2</option>
                                            <option value="1.5" ${course.grade === '1.5' ? 'selected' : ''}>1.5</option>
                                            <option value="1" ${course.grade === '1' ? 'selected' : ''}>1</option>
                                            <option value="0" ${course.grade === '0' ? 'selected' : ''}>0</option>
                                            <option value="มส." ${course.grade === 'มส.' ? 'selected' : ''}>มส.</option>
                                            <option value="ขร." ${course.grade === 'ขร.' ? 'selected' : ''}>ขร.</option>
                                            <option value="ผ." ${course.grade === 'ผ.' ? 'selected' : ''}>ผ.</option>
                                            <option value="มผ." ${course.grade === 'มผ.' ? 'selected' : ''}>มผ.</option>
                                        </select>
                                    </td>
                                    <td class="actions-cell">
                                        <button class="btn btn-info btn-sm" onclick="editCourse(${course.id}, '${semester1}')">
                                            <i class="fas fa-edit"></i>
                                        </button>
                                        <button class="btn btn-danger btn-sm" onclick="deleteCourse(${course.id}, '${semester1}')">
                                            <i class="fas fa-trash"></i>
                                        </button>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    <div class="credits-summary">
                        หน่วยกิตรวม: ${credits1}
                    </div>
                </div>
                
                <!-- Semester ${semester2} -->
                <div class="semester-column">
                    <table class="grades-table">
                        <thead>
                            <tr>
                                <th colspan="5" style="text-align: center;">ภาคเรียนที่ ${semester2}</th>
                            </tr>
                            <tr>
                                <th width="120">รหัสวิชา</th>
                                <th>ชื่อวิชา</th>
                                <th width="60">นก.</th>
                                <th width="80">เกรด</th>
                                <th width="100" class="actions-cell">การจัดการ</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${courses2.map(course => `
                                <tr>
                                    <td>${course.code}</td>
                                    <td>${course.name}</td>
                                    <td>${course.credits}</td>
                                    <td class="grade-cell">
                                        <select class="form-control grade-select" data-course-id="${course.id}" data-semester="${semester2}" style="padding: 4px; font-size: 0.8rem;">
                                            <option value="">-- เลือก --</option>
                                            <option value="4" ${course.grade === '4' ? 'selected' : ''}>4</option>
                                            <option value="3.5" ${course.grade === '3.5' ? 'selected' : ''}>3.5</option>
                                            <option value="3" ${course.grade === '3' ? 'selected' : ''}>3</option>
                                            <option value="2.5" ${course.grade === '2.5' ? 'selected' : ''}>2.5</option>
                                            <option value="2" ${course.grade === '2' ? 'selected' : ''}>2</option>
                                            <option value="1.5" ${course.grade === '1.5' ? 'selected' : ''}>1.5</option>
                                            <option value="1" ${course.grade === '1' ? 'selected' : ''}>1</option>
                                            <option value="0" ${course.grade === '0' ? 'selected' : ''}>0</option>
                                            <option value="มส." ${course.grade === 'มส.' ? 'selected' : ''}>มส.</option>
                                            <option value="ขร." ${course.grade === 'ขร.' ? 'selected' : ''}>ขร.</option>
                                            <option value="ผ." ${course.grade === 'ผ.' ? 'selected' : ''}>ผ.</option>
                                            <option value="มผ." ${course.grade === 'มผ.' ? 'selected' : ''}>มผ.</option>
                                        </select>
                                    </td>
                                    <td class="actions-cell">
                                        <button class="btn btn-info btn-sm" onclick="editCourse(${course.id}, '${semester2}')">
                                            <i class="fas fa-edit"></i>
                                        </button>
                                        <button class="btn btn-danger btn-sm" onclick="deleteCourse(${course.id}, '${semester2}')">
                                            <i class="fas fa-trash"></i>
                                        </button>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    <div class="credits-summary">
                        หน่วยกิตรวม: ${credits2}
                    </div>
                </div>
            </div>
        </div>
    `;
}

// ==================== STUDENT MANAGEMENT ====================

function showStudentModal(student = null) {
    const modalTitle = document.getElementById('studentModalTitle');
    const form = document.getElementById('studentForm');
    
    if (student) {
        modalTitle.textContent = 'แก้ไขข้อมูลนักเรียน';
        document.getElementById('studentId').value = student.id;
        document.getElementById('newStudentId').value = student.id;
        document.getElementById('newStudentName').value = student.name;
        document.getElementById('newStudentLevel').value = student.level;
        document.getElementById('newStudentMajor').value = student.major;
        document.getElementById('newStudentClass').value = student.class;
        document.getElementById('newStudentProgram').value = student.program;
        document.getElementById('newStudentYear').value = student.year || '2567';
    } else {
        modalTitle.textContent = 'เพิ่มนักเรียน';
        form.reset();
        document.getElementById('studentId').value = '';
        document.getElementById('newStudentYear').value = '2567';
        // สร้างรหัสนักเรียนใหม่อัตโนมัติ
        document.getElementById('newStudentId').value = nextStudentId++;
    }
    
    // อัพเดทตัวเลือกห้องเรียน
    updateClassOptions();
    
    studentModal.style.display = 'flex';
}

function updateClassOptions() {
    const level = document.getElementById('newStudentLevel').value;
    const major = document.getElementById('newStudentMajor').value;
    const classSelect = document.getElementById('newStudentClass');
    const programSelect = document.getElementById('newStudentProgram');
    
    // อัพเดทหลักสูตรตามระดับชั้น
    if (level.includes('ปวช.')) {
        programSelect.value = 'ปวช.';
    } else if (level.includes('ปวส.')) {
        programSelect.value = 'ปวส.';
    }
    
    // ล้างตัวเลือกเดิม
    classSelect.innerHTML = '<option value="">เลือกห้อง</option>';
    
    if (level && major) {
        const program = level.includes('ปวช.') ? 'ปวช.' : 'ปวส.';
        
        if (program === 'ปวช.') {
            // สำหรับ ปวช. ให้ใช้ชื่อสาขาเป็นชื่อห้อง
            const option = document.createElement('option');
            option.value = major;
            option.textContent = major;
            classSelect.appendChild(option);
        } else if (program === 'ปวส.') {
            // สำหรับ ปวส. ให้แสดงห้องเรียนตามสาขา
            const majorData = programData[program] && programData[program][major];
            
            if (majorData) {
                Object.keys(majorData).forEach(className => {
                    // ตรวจสอบว่าเป็นห้องเรียน (ไม่ใช่ properties อื่นๆ)
                    if (className !== 'semesters' && className !== 'courses') {
                        const option = document.createElement('option');
                        option.value = className;
                        option.textContent = className;
                        classSelect.appendChild(option);
                    }
                });
            }
        }
    }
}

async function handleStudentSubmit(e) {
    e.preventDefault();
    
    const studentId = document.getElementById('studentId').value;
    const newStudentId = document.getElementById('newStudentId').value;
    const name = document.getElementById('newStudentName').value;
    const level = document.getElementById('newStudentLevel').value;
    const major = document.getElementById('newStudentMajor').value;
    const studentClass = document.getElementById('newStudentClass').value;
    const program = document.getElementById('newStudentProgram').value;
    const year = document.getElementById('newStudentYear').value;
    const autoAddCourses = document.getElementById('autoAddCourses').checked;
    
    let result;
    
    if (studentId) {
        // แก้ไขนักเรียนที่มีอยู่ (ในระบบจริงต้องมีฟังก์ชันแก้ไข)
        showNotification('ระบบแก้ไขนักเรียนยังไม่พร้อมใช้งาน', 'warning');
        return;
    } else {
        // เพิ่มนักเรียนใหม่
        result = await callGAS('addStudent', {
            id: newStudentId,
            name: name,
            level: level,
            class: studentClass,
            program: program,
            major: major,
            year: year
        });
        
        if (result.success && autoAddCourses) {
            // เพิ่มรายวิชาตามหลักสูตร
            await loadCoursesForStudent(newStudentId, level, major, studentClass);
        }
    }
    
    if (result.success) {
        studentModal.style.display = 'none';
        showNotification('เพิ่มนักเรียนเรียบร้อยแล้ว', 'success');
        
        // โหลดข้อมูลนักเรียนใหม่
        if (currentMode === 'admin') {
            const studentResult = await callGAS('getStudentById', { id: newStudentId });
            if (studentResult.success) {
                loadStudentDataInAdminMode(studentResult.data);
            }
        }
    } else {
        showNotification('เกิดข้อผิดพลาด: ' + result.message, 'error');
    }
}

async function loadCoursesForStudent(studentId, level, major, className) {
    const program = level.includes('ปวช.') ? 'ปวช.' : 'ปวส.';
    let courseCount = 0;
    
    if (program === 'ปวช.') {
        const courses = programData[program][major]?.courses;
        if (courses) {
            for (const semester in courses) {
                for (const course of courses[semester]) {
                    const success = await addCourseToStudent(studentId, course.code, course.name, course.credits, '', semester);
                    if (success) courseCount++;
                }
            }
        }
    } else if (program === 'ปวส.') {
        const courses = programData[program][major]?.[className]?.courses;
        if (courses) {
            for (const semester in courses) {
                for (const course of courses[semester]) {
                    const success = await addCourseToStudent(studentId, course.code, course.name, course.credits, '', semester);
                    if (success) courseCount++;
                }
            }
        }
    }
    
    showNotification(`เพิ่มรายวิชา ${courseCount} รายการให้กับนักเรียนเรียบร้อยแล้ว`, 'success');
}

async function addCourseToStudent(studentId, code, name, credits, grade, semester) {
    const result = await callGAS('addCourse', {
        studentId: studentId,
        code: code,
        name: name,
        credits: credits,
        grade: grade,
        semester: semester
    });
    
    return result.success;
}

// ==================== COURSE MANAGEMENT ====================

function showCourseModal(course = null, semester = '1') {
    const modalTitle = document.getElementById('courseModalTitle');
    const form = document.getElementById('courseForm');
    
    if (course) {
        modalTitle.textContent = 'แก้ไขรายวิชา';
        document.getElementById('courseId').value = course.id;
        document.getElementById('courseCode').value = course.code;
        document.getElementById('courseName').value = course.name;
        document.getElementById('courseCredits').value = course.credits;
        document.getElementById('courseGrade').value = course.grade;
        document.getElementById('courseSemester').value = semester;
    } else {
        modalTitle.textContent = 'เพิ่มรายวิชา';
        form.reset();
        document.getElementById('courseId').value = '';
        document.getElementById('courseSemester').value = semester;
    }
    
    courseModal.style.display = 'flex';
}

async function handleCourseSubmit(e) {
    e.preventDefault();
    
    const courseId = document.getElementById('courseId').value;
    const semester = document.getElementById('courseSemester').value;
    const code = document.getElementById('courseCode').value;
    const name = document.getElementById('courseName').value;
    const credits = parseInt(document.getElementById('courseCredits').value);
    const grade = document.getElementById('courseGrade').value;
    
    let result;
    
    if (courseId) {
        // แก้ไขวิชาที่มีอยู่ (ในระบบจริงต้องมีฟังก์ชันแก้ไข)
        showNotification('ระบบแก้ไขรายวิชายังไม่พร้อมใช้งาน', 'warning');
        return;
    } else {
        // เพิ่มวิชาใหม่
        result = await callGAS('addCourse', {
            studentId: currentStudent.id,
            code: code,
            name: name,
            credits: credits,
            grade: grade,
            semester: semester
        });
    }
    
    if (result.success) {
        courseModal.style.display = 'none';
        showNotification('เพิ่มรายวิชาเรียบร้อยแล้ว', 'success');
        
        // โหลดข้อมูลใหม่
        const studentResult = await callGAS('getStudentById', { id: currentStudent.id });
        if (studentResult.success) {
            loadStudentDataInAdminMode(studentResult.data);
        }
    } else {
        showNotification('เกิดข้อผิดพลาด: ' + result.message, 'error');
    }
}

function addCourseToSemester(semester) {
    showCourseModal(null, semester);
}

function editCourse(courseId, semester) {
    const course = currentStudent.grades[semester].find(c => c.id == courseId);
    if (course) {
        showCourseModal(course, semester);
    }
}

async function deleteCourse(courseId, semester) {
    if (confirm('คุณแน่ใจว่าต้องการลบรายวิชานี้?')) {
        const result = await callGAS('deleteCourse', {
            studentId: currentStudent.id,
            courseId: courseId,
            semester: semester
        });
        
        if (result.success) {
            showNotification('ลบรายวิชาเรียบร้อยแล้ว', 'success');
            
            // โหลดข้อมูลใหม่
            const studentResult = await callGAS('getStudentById', { id: currentStudent.id });
            if (studentResult.success) {
                loadStudentDataInAdminMode(studentResult.data);
            }
        } else {
            showNotification('ลบรายวิชาไม่สำเร็จ: ' + result.message, 'error');
        }
    }
}

// ==================== BULK COURSE ADDITION ====================

function showAddCoursesModal() {
    if (!currentStudent) {
        showNotification('กรุณาเลือกนักเรียนก่อน', 'warning');
        return;
    }
    
    const modal = document.getElementById('addCoursesModal');
    const coursesList = document.getElementById('coursesList');
    const semesterSelect = document.getElementById('coursesSemester');
    
    // รีเซ็ตฟอร์ม
    coursesList.innerHTML = '';
    document.getElementById('selectedCourses').innerHTML = '';
    document.getElementById('selectedCount').textContent = '0';
    
    // โหลดรายวิชาจากหลักสูตรตามข้อมูลนักเรียน
    const program = currentStudent.program;
    const major = currentStudent.major;
    const className = currentStudent.class.trim(); // ← เพิ่ม .trim() เพื่อลบช่องว่าง
    const semester = semesterSelect.value;
    
    let availableCourses = [];
    
    console.log('Loading courses for:', { program, major, className, semester });
    
    if (program === 'ปวช.') {
        // สำหรับ ปวช.
        if (programData[program] && programData[program][major] && programData[program][major].courses) {
            availableCourses = programData[program][major].courses[semester] || [];
        }
    } else if (program === 'ปวส.') {
        // สำหรับ ปวส. - แก้ไขการดึงข้อมูลให้ถูกต้อง
        if (programData[program] && 
            programData[program][major] && 
            programData[program][major][className] && 
            programData[program][major][className].courses) {
            availableCourses = programData[program][major][className].courses[semester] || [];
        }
    }
    
    console.log('Available courses:', availableCourses);
    
    // แสดงรายวิชาที่สามารถเลือกได้
    if (availableCourses.length > 0) {
        availableCourses.forEach(course => {
            const courseItem = document.createElement('div');
            courseItem.className = 'course-item';
            courseItem.innerHTML = `
                <input type="checkbox" id="course_${course.code}" value="${course.code}" 
                       data-name="${course.name}" data-credits="${course.credits}">
                <label for="course_${course.code}">
                    <span class="course-code">${course.code}</span>
                    <span class="course-name">${course.name}</span>
                    <span class="course-credits">(${course.credits} หน่วยกิต)</span>
                </label>
            `;
            coursesList.appendChild(courseItem);
        });
        
        // เพิ่ม event listener สำหรับการเลือกวิชา
        coursesList.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
            checkbox.addEventListener('change', updateSelectedCourses);
        });
        
        modal.style.display = 'flex';
    } else {
        showNotification('ไม่พบรายวิชาในหลักสูตรสำหรับภาคเรียนนี้', 'warning');
        console.log('No courses found for:', { program, major, className, semester });
    }

// เพิ่มในฟังก์ชัน showAddCoursesModal()
console.log('=== DEBUG INFO ===');
console.log('Student:', currentStudent);
console.log('Program Data Structure:', Object.keys(programData));
console.log('Available Majors in ปวส.:', programData['ปวส.'] ? Object.keys(programData['ปวส.']) : 'No data');
console.log('Available Classes in Major:', programData['ปวส.'] && programData['ปวส.'][currentStudent.major] ? Object.keys(programData['ปวส.'][currentStudent.major]) : 'No data');

}

function updateSelectedCourses() {
    const selectedCoursesDiv = document.getElementById('selectedCourses');
    const selectedCountSpan = document.getElementById('selectedCount');
    const checkboxes = document.querySelectorAll('#coursesList input[type="checkbox"]:checked');
    
    selectedCoursesDiv.innerHTML = '';
    selectedCountSpan.textContent = checkboxes.length;
    
    checkboxes.forEach(checkbox => {
        const courseDiv = document.createElement('div');
        courseDiv.className = 'selected-course-item';
        courseDiv.innerHTML = `
            <span class="course-code">${checkbox.value}</span>
            <span class="course-name">${checkbox.dataset.name}</span>
            <span class="course-credits">(${checkbox.dataset.credits} หน่วยกิต)</span>
        `;
        selectedCoursesDiv.appendChild(courseDiv);
    });
}

async function confirmAddCourses() {
    const semester = document.getElementById('coursesSemester').value;
    const checkboxes = document.querySelectorAll('#coursesList input[type="checkbox"]:checked');
    
    if (checkboxes.length === 0) {
        showNotification('กรุณาเลือกรายวิชาอย่างน้อยหนึ่งรายวิชา', 'warning');
        return;
    }
    
    let successCount = 0;
    
    for (const checkbox of checkboxes) {
        const result = await callGAS('addCourse', {
            studentId: currentStudent.id,
            code: checkbox.value,
            name: checkbox.dataset.name,
            credits: checkbox.dataset.credits,
            grade: '',
            semester: semester
        });
        
        if (result.success) {
            successCount++;
        }
    }
    
    document.getElementById('addCoursesModal').style.display = 'none';
    showNotification(`เพิ่มรายวิชาเรียบร้อย ${successCount} รายการ`, 'success');
    
    // โหลดข้อมูลนักเรียนใหม่
    const studentResult = await callGAS('getStudentById', { id: currentStudent.id });
    if (studentResult.success) {
        loadStudentDataInAdminMode(studentResult.data);
    }
}

// ==================== IMPORT/EXPORT ====================

function showImportExportModal(tab) {
    document.getElementById('importExportTitle').textContent = tab === 'import' ? 'นำเข้าข้อมูล' : 'ส่งออกข้อมูล';
    switchImportExportTab(tab);
    importExportModal.style.display = 'flex';
}

function switchImportExportTab(tabName) {
    // อัพเดท tabs
    document.querySelectorAll('.modal-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelector(`.modal-tab[data-tab="${tabName}"]`).classList.add('active');
    
    // อัพเดท content
    document.querySelectorAll('.modal-tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`${tabName}Tab`).classList.add('active');
}

function importData() {
    const fileInput = document.getElementById('importFile');
    const importType = document.getElementById('importType').value;
    
    if (!fileInput.files[0]) {
        showNotification('กรุณาเลือกไฟล์', 'warning');
        return;
    }
    
    const file = fileInput.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const importedData = JSON.parse(e.target.result);
            showNotification('นำเข้าข้อมูลเรียบร้อยแล้ว (จำลอง)', 'success');
            importExportModal.style.display = 'none';
            
        } catch (error) {
            showNotification('ไฟล์ไม่ถูกต้อง', 'error');
        }
    };
    
    reader.readAsText(file);
}

function exportData() {
    const exportType = document.getElementById('exportType').value;
    const exportFormat = document.getElementById('exportFormat').value;
    
    if (exportFormat === 'csv') {
        exportToCSV(exportType);
    } else {
        // JSON export (existing code)
        let dataToExport = {};
        
        if (exportType === 'all') {
            dataToExport = { message: "การส่งออกข้อมูลทั้งหมด" };
        } else if (exportType === 'students') {
            dataToExport = { message: "การส่งออกข้อมูลนักเรียน" };
        } else if (exportType === 'courses') {
            dataToExport = { message: "การส่งออกข้อมูลรายวิชา" };
        } else if (exportType === 'grades') {
            dataToExport = { message: "การส่งออกข้อมูลเกรด" };
        }
        
        const dataStr = JSON.stringify(dataToExport, null, 2);
        const dataBlob = new Blob([dataStr], { type: 'application/json' });
        downloadFile(dataBlob, `student_data_${new Date().toISOString().split('T')[0]}.json`);
    }
    
    showNotification('ส่งออกข้อมูลเรียบร้อยแล้ว', 'success');
    importExportModal.style.display = 'none';
}

async function exportToCSV(exportType) {
    try {
        let csvContent = '';
        const result = await callGAS('getAllData');
        
        if (!result.success) {
            showNotification('ไม่สามารถดึงข้อมูลได้: ' + result.message, 'error');
            return;
        }
        
        const data = result.data;
        
        if (exportType === 'students' || exportType === 'all') {
            // สร้าง CSV สำหรับข้อมูลนักเรียน
            csvContent += "ข้อมูลนักเรียน\n";
            csvContent += "รหัสนักเรียน,ชื่อ-สกุล,ระดับชั้น,ห้อง,หลักสูตร,สาขาวิชา,ปีการศึกษา\n";
            
            data.students.forEach(student => {
                csvContent += `"${student.id}","${student.name}","${student.level}","${student.class}","${student.program}","${student.major}","${student.year}"\n`;
            });
            
            csvContent += "\n\n";
        }
        
        if (exportType === 'courses' || exportType === 'all') {
            // สร้าง CSV สำหรับข้อมูลรายวิชา
            csvContent += "ข้อมูลรายวิชา\n";
            csvContent += "รหัสรายวิชา,รหัสนักเรียน,ชื่อวิชา,หน่วยกิต,เกรด,ภาคเรียน\n";
            
            data.courses.forEach(course => {
                csvContent += `"${course.code}","${course.student_id}","${course.name}",${course.credits},"${course.grade}","${course.semester}"\n`;
            });
        }
        
        // สร้าง BLOB พร้อม BOM สำหรับรองรับภาษาไทยใน Excel
        const BOM = '\uFEFF';
        const csvBlob = new Blob([BOM + csvContent], { 
            type: 'text/csv;charset=utf-8;' 
        });
        
        downloadFile(csvBlob, `student_data_${new Date().toISOString().split('T')[0]}.csv`);
        
    } catch (error) {
        console.error('Error exporting to CSV:', error);
        showNotification('เกิดข้อผิดพลาดในการส่งออกข้อมูล: ' + error.message, 'error');
    }
}

function downloadFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ==================== GRADE MANAGEMENT ====================

async function updateGradeInDatabase(courseId, semester, grade) {
    const result = await callGAS('updateGrade', {
        studentId: currentStudent.id,
        courseId: courseId,
        semester: semester,
        grade: grade
    });
    
    if (result.success) {
        showNotification('อัพเดทเกรดเรียบร้อยแล้ว', 'success');
        
        // อัพเดท GPA
        const gpaValue = document.querySelector('.gpa-value');
        const studentResult = await callGAS('getStudentById', { id: currentStudent.id });
        if (studentResult.success) {
            gpaValue.textContent = calculateGPA(studentResult.data).toFixed(2);
        }
    } else {
        showNotification('อัพเดทเกรดไม่สำเร็จ: ' + result.message, 'error');
    }
}

function saveAllGrades() {
    // ในระบบจริงจะบันทึกลงฐานข้อมูล
    showNotification('บันทึกข้อมูลทั้งหมดเรียบร้อยแล้ว', 'success');
}

// ==================== FILTERS ====================

function applyFilters() {
    // ในระบบจริงจะกรองข้อมูลตามที่เลือก
    showNotification('ใช้การกรองข้อมูลแล้ว', 'info');
}

// ==================== UTILITY FUNCTIONS ====================

function calculateGPA(student) {
    let totalCredits = 0;
    let totalPoints = 0;
    
    for (const semester in student.grades) {
        student.grades[semester].forEach(course => {
            const gradeValue = parseFloat(course.grade);
            if (!isNaN(gradeValue)) {
                totalCredits += course.credits;
                totalPoints += course.credits * gradeValue;
            }
        });
    }
    
    return totalCredits > 0 ? totalPoints / totalCredits : 0;
}

function calculateCredits(courses) {
    return courses.reduce((sum, course) => sum + course.credits, 0);
}

// ==================== NOTIFICATION SYSTEM ====================

function showNotification(message, type = 'info') {
    // สร้าง notification element
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <div class="notification-content">
            <span class="notification-message">${message}</span>
            <button class="notification-close">&times;</button>
        </div>
    `;
    
    // สไตล์ notification
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : type === 'warning' ? '#f59e0b' : '#3b82f6'};
        color: white;
        padding: 12px 16px;
        border-radius: 6px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        z-index: 1000;
        max-width: 350px;
        animation: slideIn 0.3s ease;
    `;
    
    document.body.appendChild(notification);
    
    // ปิด notification อัตโนมัติ
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 300);
    }, 3000);
    
    // ปิดเมื่อคลิก
    notification.querySelector('.notification-close').addEventListener('click', () => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 300);
    });
}

// ==================== GAS COMMUNICATION ====================

async function callGAS(action, parameters = {}) {
    try {
        // สร้าง URL parameters
        const urlParams = new URLSearchParams();
        urlParams.append('action', action);
        
        Object.keys(parameters).forEach(key => {
            if (parameters[key] !== undefined && parameters[key] !== null) {
                urlParams.append(key, parameters[key]);
            }
        });
        
        const url = GAS_URL + '?' + urlParams.toString();
        console.log('Calling GAS:', url);
        
        const response = await fetch(url, {
            method: 'GET',
            mode: 'cors'
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        return result;
        
    } catch (error) {
        console.error('Error calling GAS:', error);
        return { 
            success: false, 
            message: 'การเชื่อมต่อล้มเหลว: ' + error.message 
        };
    }
}

// ==================== CSS INJECTION ====================

// เพิ่ม CSS สำหรับ animation
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    
    @keyframes slideOut {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
    }
    
    .notification-close {
        background: none;
        border: none;
        color: white;
        font-size: 1.1rem;
        cursor: pointer;
        margin-left: 8px;
    }
    
    .notification-content {
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    
    .form-check {
        display: flex;
        align-items: center;
        margin-bottom: 1rem;
    }
    
    .form-check-input {
        margin-right: 8px;
    }
    
    .form-check-label {
        margin-bottom: 0;
        font-weight: normal;
    }
    
    /* สไตล์สำหรับ modal เพิ่มรายวิชาจากหลักสูตร */
    .courses-list-container {
        max-height: 300px;
        overflow-y: auto;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        padding: 10px;
        margin-bottom: 1rem;
    }
    
    .selected-courses-summary {
        border: 1px solid var(--border-color);
        border-radius: 4px;
        padding: 10px;
        background: var(--bg-tertiary);
    }
    
    .selected-courses {
        max-height: 150px;
        overflow-y: auto;
    }
    
    .selected-course-item {
        padding: 8px;
        border-bottom: 1px solid var(--border-color);
        display: flex;
        gap: 10px;
        align-items: center;
    }
    
    .selected-course-item:last-child {
        border-bottom: none;
    }
    
    .modal-footer {
        display: flex;
        justify-content: flex-end;
        gap: 10px;
        margin-top: 1rem;
        padding-top: 1rem;
        border-top: 1px solid var(--border-color);
    }
    
    /* สไตล์สำหรับรายวิชาที่เลือก */
    .course-item {
        display: flex;
        align-items: center;
        padding: 8px 0;
        border-bottom: 1px solid var(--border-color);
    }
    
    .course-item:last-child {
        border-bottom: none;
    }
    
    .course-item input[type="checkbox"] {
        margin-right: 10px;
    }
    
    .course-item label {
        flex: 1;
        cursor: pointer;
        margin-bottom: 0;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .course-code {
        font-weight: bold;
        color: var(--primary);
        min-width: 120px;
    }
    
    .course-name {
        flex: 1;
    }
    
    .course-credits {
        color: var(--text-secondary);
        font-size: 0.8rem;
    }
`;
document.head.appendChild(style);