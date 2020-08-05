package oracle.dao;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class OracleDao {

	/**
	 * 整理No単位のデータ取得
	 * @param centerCode センターコード
	 * @param nyukoYoteiDt 入庫予定日
	 * @return 整理No単位のマップ（Key：整理No）
	 * @throws KENPINException DBアクセスエラー発生
	 */
	public String excuteSQL(String url, String user, String pwd, String sql) throws Exception {

		StringBuffer sb = new StringBuffer();

		Connection con = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;

		con = connect(url, user, pwd);
		if(null==con){
			throw new Exception(String.format("connect failed. url:%s\r\nuser:%s\r\npwd:%s", url, user, pwd));
		}

		try {
			//「検品済」レコードの件数取得
			pstmt = con.prepareStatement(sql);
			rs = pstmt.executeQuery();

			for(int columnIndex = 0; columnIndex < rs.getMetaData().getColumnCount(); columnIndex++) {
				sb.append(rs.getMetaData().getColumnName(columnIndex + 1));
				if (columnIndex < rs.getMetaData().getColumnCount() - 1) {
					sb.append(",");
				} else {
					sb.append("\r\n");
				}
			}

			while(rs.next()){
				for(int columnIndex = 0; columnIndex < rs.getMetaData().getColumnCount(); columnIndex++) {
					sb.append(rs.getString(columnIndex + 1));
					if (columnIndex < rs.getMetaData().getColumnCount() - 1) {
						sb.append(",");
					} else {
						sb.append("\r\n");
					}
				}
	        }


		} catch (Exception e) {
			e.printStackTrace();
			sb = new StringBuffer();
		}finally {
			try {
				if(rs != null)rs.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
			try {
				if(pstmt != null) pstmt.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
			try {
	            if(con != null) con.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}

		return sb.toString();

	}

	/**
	 * Lixxiデータベース接続
	 * @return コネクション
	 */
	private Connection connect(String url, String user, String pwd) {
		Connection con = null;
		try {
			con = DriverManager.getConnection(url,
					user, //user
					pwd); //password;

			return con;
		} catch (Exception e) {
			e.printStackTrace();
			con = null;
		}
		return con;

	}
}
