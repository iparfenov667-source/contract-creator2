import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";

interface ContractData {
  contract_num: string;
  date_zakl: string;
  fio: string;
  data_rod: string;
  passport: string;
  summa: number;
  sut_por: string;
  adres: string;
  phone: string;
  payment: number;
  months: number;
  nds: number;
}

function generatePaymentDates(dateStr: string, months: number): string[] {
  const parts = dateStr.split(".");
  if (parts.length !== 3) return [];
  
  const day = parseInt(parts[0]);
  const month = parseInt(parts[1]) - 1;
  const year = parseInt(parts[2]);
  const startDate = new Date(year, month, day);
  
  const dates: string[] = [];
  for (let i = 0; i < months; i++) {
    if (i === 0) {
      dates.push(dateStr);
    } else {
      const newMonth = (startDate.getMonth() + i) % 12;
      const newYear = startDate.getFullYear() + Math.floor((startDate.getMonth() + i) / 12);
      dates.push(`10.${String(newMonth + 1).padStart(2, "0")}.${newYear}`);
    }
  }
  return dates;
}

export async function generateContract(data: ContractData): Promise<void> {
  const response = await fetch("/Dogovor_OBSHAYA_ShABLON.docx");
  const arrayBuffer = await response.arrayBuffer();
  
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
  });

  const paymentDates = generatePaymentDates(data.date_zakl, data.months);

  const context: Record<string, string | number> = {
    contract_num: data.contract_num,
    date_zakl: data.date_zakl,
    fio: data.fio,
    data_rod: data.data_rod,
    passport: data.passport,
    summa: data.summa.toLocaleString("ru-RU"),
    sut_por: data.sut_por,
    srok: data.months,
    adres: data.adres,
    phone: data.phone,
  };

  for (let i = 1; i <= 24; i++) {
    context[`payment_date_${i}`] = i <= data.months ? paymentDates[i - 1] : "";
    context[`payment_summa_${i}`] = i <= data.months ? data.payment.toLocaleString("ru-RU") : "";
  }

  doc.render(context);

  const out = doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  const filename = `${data.contract_num.replace("№", "").trim()}.docx`;
  saveAs(out, filename);
}
