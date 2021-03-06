package oracle.dao;


public class DaoUtil {

	public static void main(String[] args) {
		OracleDao dao = new OracleDao();
		try {
			if (args.length < 4) {
				return;
			}
			/*
			String str = dao.excuteSQL("jdbc:oracle:thin:@10.35.21.50:1551:sthokan",
					"ST_HOKAN", "ST_HOKAN44",
					"SELECT HI1.CENTERCD, HI1.NYUKOYOTEIDT , HI1.SEIRINO, count(HI1.KURAIRENO) KURAIRECOUNT, sum(HI1.NYUKOYOTEIKOSU) KOSUSUM, sum(HI1.NYUKOYOTEIHASU) HASUSUM, sum(CASE WHEN HI1.JURYOTANI = '2' then (HI1.NYUKOYOTEIJURYO *  0.45359) ELSE HI1.NYUKOYOTEIJURYO END) JURYOSUM, HY1.ATSUKAIBINNM, HY1.SHABAN, HY1.DRIVERNM FROM HINYUKOYOTEI HI1, (select CENTERCD, SEIRINO, ATSUKAIBINNM, SHABAN, DRIVERNM  from HYTOCHAKUSHARYO  where NVL(HYTOCHAKUSHARYO.TORIKESHIFG, '0') = '0'  group by CENTERCD, SEIRINO, ATSUKAIBINNM, SHABAN, DRIVERNM  ) HY1 WHERE AKAKUROKB = '0' AND NVL(HI1.TORIKESHIFG, '0') = '0' AND NOT TRANCD IN ('NJ01', 'NR01') AND NOT (TRANCD = 'NS01' AND SUBSTR(NVL(HI1.DATAKB, '000'), 2, 1) = '2') AND HI1.CENTERCD = HY1.CENTERCD AND HI1.SEIRINO = HY1.SEIRINO GROUP BY (HI1.CENTERCD, HI1.NYUKOYOTEIDT, HI1.SEIRINO, HI1.NYUKOYOTEIDT, HY1.ATSUKAIBINNM, HY1.DRIVERNM, HY1.SHABAN) ORDER BY HI1.SEIRINO");
					*/
			String str = dao.excuteSQL(args[0],
					args[1],
					args[2],
					args[3]);
			System.out.print(str);
		} catch (Exception e) {
			// TODO 自動生成された catch ブロック
			e.printStackTrace();
		}

	}

}
