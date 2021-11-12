import XLSX from "xlsx";

const EXCLUDE_KEYS = ["!ref", "!merges", "!margins"];

const subtractLetter = (value) => {
  //* A1 A2 AA3 AB2 AB3 AAAAB1

  //* AAAAB1 [A,A,A,A,B,1,2,3] => [A,A,A,A,B] => AAAB
  const regex = /^[a-zA-Z]+$/;
  return value
    .split("")
    .filter((item) => regex.test(item))
    .join("");
};

const subtractNumber = (value) => {
  const regex = /^[0-9]*$/;
  return Number(
    value
      .split("")
      .filter((item) => regex.test(item))
      .join("")
  );
};

const generateCoordinate = ({ letters, value }) => {
  return Object.keys(letters).map(
    (letter) => `${letter.toUpperCase()}${value}`
  );
};

const generateRow = ({ headers }) => {
  const row: any = {};
  Object.values(headers).forEach((value: any) => {
    row[value] = "";
  });

  return row;
};

const getLetters = ({ sheet }: any) => {
  return new Promise((resolve, reject) => {
    const letters = {};

    Object.keys(sheet).forEach((key) => {
      const letter = subtractLetter(key);
      if (!letters[letter] && !EXCLUDE_KEYS.includes(key)) {
        letters[letter] = letter;
      }
    });

    resolve({ letters });
  });
};

const getRange = ({ sheet, letters }: any) => {
  return new Promise((resolve, reject) => {
    const letter = Object.keys(letters)[0];
    const values = [letter.toUpperCase()];

    const res = Object.keys(sheet)
      .filter((key) => values.includes(subtractLetter(key)))
      .map((item) => subtractNumber(item))
      .sort((a, b) => a - b);

    resolve({
      min: res[0],
      max: res[res.length - 1],
      values: res,
    });
  });
};

const getHeaders = ({ sheet, letters, min }: any) => {
  return new Promise((resolve, reject) => {
    const headersKeys = generateCoordinate({ letters, value: min });
    const headers: any = {};

    Object.keys(sheet).forEach((key) => {
      if (headersKeys.includes(key)) {
        const letter = subtractLetter(key);
        headers[letter] = sheet[key].v;
      }
    });

    resolve({ headers });
  });
};

const getRow = ({ headers, sheet, letters, num }) => {
  return new Promise((resolve, reject) => {
    const keys: any[] = generateCoordinate({ letters, value: num });
    const row: any = generateRow({ headers });
    // const row = {};
    Object.keys(sheet).forEach((key) => {
      if (keys.includes(key)) {
        const header = headers[subtractLetter(key)];
        row[header] = !sheet[key].v ? "" : sheet[key].v;
      }
    });

    resolve({ row });
  });
};

//eslint-disable-next-line
const getSheets = async ({ SHEETS }) => {
  return SHEETS.map(async (sheet) => {
    const { letters }: any = await getLetters({ sheet });
    const { min, values }: any = await getRange({
      sheet,
      letters,
    });
    const { headers }: any = await getHeaders({
      sheet,
      letters,
      min,
    });

    const res = values
      .filter((num) => num !== min)
      .map(async (num) => {
        const { row }: any = await getRow({
          headers,
          sheet,
          letters,
          num,
        });

        return row;
      });

    const rows = await Promise.all(res);
    // console.log(min);
    // console.log(values, "values");
    // console.log(letters, "letters");
    // console.log(headers);
    // console.log(rows, "este es mi resultado XD");

    return rows;
  });
};

export async function handleDropAsync(e) {
  e.stopPropagation();
  e.preventDefault();
  const f = e.dataTransfer.files[0];
  const data = await f.arrayBuffer();
  const workbook = XLSX.read(data);

  const SHEETS = Object.values(workbook.Sheets).filter(
    (sheet) => sheet["!ref"]
  );

  const sheets = await getSheets({ SHEETS });

  const res = await Promise.all(sheets);

  console.log(res.length);
  console.log(res);
}
