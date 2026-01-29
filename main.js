const path = require('path');
const fs = require('fs-extra');
const util = require('util');
const xlsx = require('xlsx');

const writeFile = util.promisify(fs.writeFile).bind(util);

// ファイル設定
const src_file = path.resolve(path.join('xlsx', 'sheet.xlsx')); //エクセルファイル名
const sheet_name = '2026'; //シート名
const dest_file = path.resolve(path.join('dest', 'list.json')); //jsonファイル名
const dest_lp_file = path.resolve(path.join('dest', 'list-under.json')); //下層用jsonファイル名

// エクセル範囲の設定
const EXCEL_START_ROW = 2;  // データ開始行
const EXCEL_START_ROW_INDEX = EXCEL_START_ROW - 1;
const EXCEL_END_ROW = 38;   // データ終了行
const EXCEL_RANGE = `A${EXCEL_START_ROW}:M${EXCEL_END_ROW}`;  // エクセル範囲

const readXLSX = () => {
  let ret = [];

  try {
    const book = xlsx.readFile(src_file);
    const sheet = book.Sheets[sheet_name];

    if (!sheet) {
      throw new Error(`シート "${sheet_name}" が見つかりませんでした。`);
    }

    sheet['!ref'] = EXCEL_RANGE;

    // セルの値を取得する関数（空の場合は 'none'）
    const getCellValue = (col, row, defaultValue = 'none') => {
      const cell = sheet[xlsx.utils.encode_cell({ c: col, r: row })];
      if (!cell || cell.v === undefined || cell.v.toString().trim() === '') {
        return defaultValue;
      }
      return cell.v.toString().trim();
    };

    for(let r = EXCEL_START_ROW_INDEX; r < EXCEL_END_ROW; r++) {
      // 基本情報の取得
      const r_id = getCellValue(0, r);
      const r_name = getCellValue(1, r);
      const r_des = getCellValue(2, r);
      const r_link2 = getCellValue(3, r);
      const r_link3 = getCellValue(4, r);
      const r_link1 = getCellValue(9, r);
      const r_link4_raw = getCellValue(10, r);
      const r_link4 = r_link4_raw === '-' ? 'none' : r_link4_raw;
      const r_photo = getCellValue(11, r);
      const r_entry = getCellValue(12, r);
      const r_category = getCellValue(13, r);

      // リスト形式のデータを取得
      const getListValue = (col) => {
        const cell = sheet[xlsx.utils.encode_cell({ c: col, r: r })];
        if (!cell || !cell.v) return ["すべて"];
        return ["すべて", ...cell.v.toString().trim().split('\r\n')];
      };

      const r_syokusyus = getListValue(5);
      const r_gakubus = getListValue(6);
      const r_bunyas = getListValue(7);
      const r_kodawaris = getListValue(8);

      ret.push({
        id: r_id,
        name: r_name,
        des: r_des,
        photo: r_photo,
        syokusyus: r_syokusyus,
        gakubus: r_gakubus,
        bunyas: r_bunyas,
        kodawaris: r_kodawaris,
        link1: r_link1,
        link2: r_link2,
        link3: r_link3,
        link4: r_link4,
        entry: r_entry,
        category: r_category,
      });
    }

    return ret;
  } catch (error) {
    console.error('Excelファイルの読み込みに失敗しました:', error);
    throw error;
  }
};

const createJSON = async (data) => {
  try {
    // データを処理
    const processedData = data.map(item => {
      // 配列内の文字列を分割する関数
      const splitStrings = (arr) => {
        return arr.reduce((result, str) => {
          if (str === 'すべて') {
            result.push(str);
          } else {
            // \n で分割して配列に追加
            result.push(...str.split('\n'));
          }
          return result;
        }, []);
      };

      const { entry, category, ...rest } = item;
      return {
        ...rest,
        syokusyus: splitStrings(item.syokusyus),
        gakubus: splitStrings(item.gakubus),
        bunyas: splitStrings(item.bunyas),
        kodawaris: splitStrings(item.kodawaris)
      };
    });

    // list.json も下層用と同じく items 配列でラップする
    const json = JSON.stringify({ items: processedData }, null, '\t');
    await writeFile(dest_file, json);
    console.log('JSONファイル（list.json）の作成に成功しました。');
  } catch (error) {
    console.error('JSONファイル（list.json）の作成に失敗しました。:', error);
    throw error;
  }
};

const createLPJSON = async (data) => {
  try {
    // 下層用のデータを処理
    const lpData = {
      items: data.map(item => ({
        id: item.id,
        name: item.name,
        photo: item.photo,
        entry: item.entry,
        category: item.category
      }))
    };

    const json = JSON.stringify(lpData, null, '\t');
    await writeFile(dest_lp_file, json);
    console.log('下層用JSONファイル（list-under.json）の作成に成功しました。');
  } catch (error) {
    console.error('下層用JSONファイル（list-under.json）の作成に失敗しました:', error);
    throw error;
  }
};

// メイン処理
(async () => {
  try {
    const data = readXLSX();
    await createJSON(data);
    await createLPJSON(data);
  } catch (error) {
    console.error('ファイル処理に失敗しました:', error);
    process.exit(1);
  }
})();
