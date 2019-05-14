Attribute VB_Name = "Module1"
Option Explicit

Function wenkuwikiURL(sname As String, newbookurl As String) As String

Select Case sname
Case "讲谈社X文库White Heart"
wenkuwikiURL = "講談社X文庫ホワイトハート"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E8%AC%9B%E8%AB%87%E7%A4%BEX%E6%96%87%E5%BA%AB%E3%83%9B%E3%83%AF%E3%82%A4%E3%83%88%E3%83%8F%E3%83%BC%E3%83%88%2Cp_lbr_publishers_browse-bin%3A%E8%AC%9B%E8%AB%87%E7%A4%BE&sort=date-desc-rank&keywords=%E8%AC%9B%E8%AB%87%E7%A4%BEX%E6%96%87%E5%BA%AB%E3%83%9B%E3%83%AF%E3%82%A4%E3%83%88%E3%83%8F%E3%83%BC%E3%83%88&ie=UTF8&qid=1436578007&lo=digital-text"

Case "Cobalt文库"
wenkuwikiURL = "コバルト文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%82%B3%E3%83%90%E3%83%AB%E3%83%88%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E9%9B%86%E8%8B%B1%E7%A4%BE&sort=date-desc-rank&keywords=%E3%82%B3%E3%83%90%E3%83%AB%E3%83%88%E6%96%87%E5%BA%AB&ie=UTF8&qid=1436578011&lo=digital-text"

Case "角川Beans文库"
wenkuwikiURL = "角川ビーンズ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E8%A7%92%E5%B7%9D%E3%83%93%E3%83%BC%E3%83%B3%E3%82%BA%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E8%A7%92%E5%B7%9D%E6%9B%B8%E5%BA%97&sort=date-desc-rank&keywords=%E8%A7%92%E5%B7%9D%E3%83%93%E3%83%BC%E3%83%B3%E3%82%BA%E6%96%87%E5%BA%AB&ie=UTF8&qid=1436578824&lo=digital-text"

Case "Bs-Log文库"
wenkuwikiURL = "ビーズログ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%83%93%E3%83%BC%E3%82%BA%E3%83%AD%E3%82%B0%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E3%82%A8%E3%83%B3%E3%82%BF%E3%83%BC%E3%83%96%E3%83%AC%E3%82%A4%E3%83%B3&sort=relevancerank&keywords=%E3%83%93%E3%83%BC%E3%82%BA%E3%83%AD%E3%82%B0%E6%96%87%E5%BA%AB&ie=UTF8&qid=1436578015&lo=digital-text"

Case "LuLuLu文库"
wenkuwikiURL = "ルルル文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%83%AB%E3%83%AB%E3%83%AB%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E5%B0%8F%E5%AD%A6%E9%A4%A8&sort=date-desc-rank&keywords=%E3%83%AB%E3%83%AB%E3%83%AB%E6%96%87%E5%BA%AB&ie=UTF8&qid=1436579132&lo=digital-text"

Case "一迅社文库Iris"
wenkuwikiURL = "一迅社文庫アイリス"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E4%B8%80%E8%BF%85%E7%A4%BE%E6%96%87%E5%BA%AB%E3%82%A2%E3%82%A4%E3%83%AA%E3%82%B9%2Cp_lbr_publishers_browse-bin%3A%E4%B8%80%E8%BF%85%E7%A4%BE&sort=date-desc-rank&keywords=%E4%B8%80%E8%BF%85%E7%A4%BE%E6%96%87%E5%BA%AB%E3%82%A2%E3%82%A4%E3%83%AA%E3%82%B9&ie=UTF8&qid=1436579135&lo=digital-text"

Case "头饰文库(tiara)"
wenkuwikiURL = "ティアラ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&fst=as%3Aoff&keywords=%E3%83%86%E3%82%A3%E3%82%A2%E3%83%A9%E6%96%87%E5%BA%AB&qid=1432471833&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%83%86%E3%82%A3%E3%82%A2%E3%83%A9%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E3%83%97%E3%83%A9%E3%83%B3%E3%82%BF%E3%83%B3%E5%87%BA%E7%89%88&sort=date-desc-rank"

Case "香草文库(Vanilla)"
wenkuwikiURL = "ヴァニラ文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&ie=UTF8&keywords=%22%E3%83%B4%E3%82%A1%E3%83%8B%E3%83%A9%E6%96%87%E5%BA%AB%22&lo=digital-text&qid=1435727358&rh=k%3A%22%E3%83%B4%E3%82%A1%E3%83%8B%E3%83%A9%E6%96%87%E5%BA%AB%22%2Cn%3A2250738051&sort=date-desc-rank"

Case "TL蜜姫文库"
wenkuwikiURL = "TL◆蜜姫文庫チュチュ"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=TL%E2%97%86%E8%9C%9C%E5%A7%AB%E6%96%87%E5%BA%AB%E3%83%81%E3%83%A5%E3%83%81%E3%83%A5&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3ATL%E2%97%86%E8%9C%9C%E5%A7%AB%E6%96%87%E5%BA%AB%E3%83%81%E3%83%A5%E3%83%81%E3%83%A5&qid=1446987511&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"

Case "糖果巧克力文库"
wenkuwikiURL = "TL ボンボンショコラ文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3ATL+%E3%83%9C%E3%83%B3%E3%83%9C%E3%83%B3%E3%82%B7%E3%83%A7%E3%82%B3%E3%83%A9%E6%96%87%E5%BA%AB&sort=date-desc-rank&keywords=TL+%E3%83%9C%E3%83%B3%E3%83%9C%E3%83%B3%E3%82%B7%E3%83%A7%E3%82%B3%E3%83%A9%E6%96%87%E5%BA%AB&ie=UTF8&qid=1446986852&lo=digital-text"

Case "zh"
wenkuwikiURL = "中文版"
newbookurl = "http://www.wenku8.cn/modules/article/articlelist.php?class=12"

Case "富士见Fantasia文库"
wenkuwikiURL = "富士見ファンタジア文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E5%AF%8C%E5%A3%AB%E8%A6%8B%E3%83%95%E3%82%A1%E3%83%B3%E3%82%BF%E3%82%B8%E3%82%A2%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E5%AF%8C%E5%A3%AB%E8%A6%8B%E6%9B%B8%E6%88%BF&sort=date-desc-rank&keywords=%E5%AF%8C%E5%A3%AB%E8%A6%8B%E3%83%95%E3%82%A1%E3%83%B3%E3%82%BF%E3%82%B8%E3%82%A2%E6%96%87%E5%BA%AB&ie=UTF8&qid=1439632274&lo=digital-text"
Case "MF文库J"
wenkuwikiURL = "MF文庫J"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_relevancerank?lo=digital-text&keywords=MF%E6%96%87%E5%BA%ABJ&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3AMF%E6%96%87%E5%BA%ABJ%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E3%83%A1%E3%83%87%E3%82%A3%E3%82%A2%E3%83%95%E3%82%A1%E3%82%AF%E3%83%88%E3%83%AA%E3%83%BC&qid=1440077484&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=relevancerank"
Case "电击文库"
wenkuwikiURL = "電撃文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E9%9B%BB%E6%92%83%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E3%82%A2%E3%82%B9%E3%82%AD%E3%83%BC%E3%83%BB%E3%83%A1%E3%83%87%E3%82%A3%E3%82%A2%E3%83%AF%E3%83%BC%E3%82%AF%E3%82%B9&keywords=%E9%9B%BB%E6%92%83%E6%96%87%E5%BA%AB&ie=UTF8&qid=1439630591&lo=digital-text"
Case "角川Sneaker文库"
wenkuwikiURL = "角川スニーカー文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E8%A7%92%E5%B7%9D%E3%82%B9%E3%83%8B%E3%83%BC%E3%82%AB%E3%83%BC%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E8%A7%92%E5%B7%9D%E6%9B%B8%E5%BA%97&keywords=%E8%A7%92%E5%B7%9D%E3%82%B9%E3%83%8B%E3%83%BC%E3%82%AB%E3%83%BC%E6%96%87%E5%BA%AB&ie=UTF8&qid=1439631925&lo=digital-text"
Case "GA文库"
wenkuwikiURL = "GA文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=GA%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3AGA%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3ASB%E3%82%AF%E3%83%AA%E3%82%A8%E3%82%A4%E3%83%86%E3%82%A3%E3%83%96&qid=1439631858&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"
Case "Fami通文库"
wenkuwikiURL = "ファミ通文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_relevancerank?lo=digital-text&keywords=%E3%83%95%E3%82%A1%E3%83%9F%E9%80%9A%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%83%95%E3%82%A1%E3%83%9F%E9%80%9A%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E3%82%A8%E3%83%B3%E3%82%BF%E3%83%BC%E3%83%96%E3%83%AC%E3%82%A4%E3%83%B3&qid=1440077655&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=relevancerank"
Case "Gagaga文库"
wenkuwikiURL = "ガガガ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E3%82%AC%E3%82%AC%E3%82%AC%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%82%AC%E3%82%AC%E3%82%AC%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E5%B0%8F%E5%AD%A6%E9%A4%A8&qid=1439631204&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"
Case "讲谈社轻小说文库"
wenkuwikiURL = "講談社ラノベ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E8%AC%9B%E8%AB%87%E7%A4%BE%E3%83%A9%E3%83%8E%E3%83%99%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E8%AC%9B%E8%AB%87%E7%A4%BE%E3%83%A9%E3%83%8E%E3%83%99%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E8%AC%9B%E8%AB%87%E7%A4%BE&qid=1439631580&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"
Case "集英社SuperDash文库" To "DashX文库"
wenkuwikiURL = "集英社スーパーダッシュ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E9%9B%86%E8%8B%B1%E7%A4%BE+%E3%82%B9%E3%83%BC%E3%83%91%E3%83%BC%E3%83%80%E3%83%83%E3%82%B7%E3%83%A5%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E9%9B%86%E8%8B%B1%E7%A4%BE+%E3%82%B9%E3%83%BC%E3%83%91%E3%83%BC%E3%83%80%E3%83%83%E3%82%B7%E3%83%A5%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E9%9B%86%E8%8B%B1%E7%A4%BE&qid=1439631378&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"
Case "一迅社文库"
wenkuwikiURL = "一迅社文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E4%B8%80%E8%BF%85%E7%A4%BE%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E4%B8%80%E8%BF%85%E7%A4%BE&sort=date-desc-rank&keywords=%E4%B8%80%E8%BF%85%E7%A4%BE%E6%96%87%E5%BA%AB&ie=UTF8&qid=1439631363&lo=digital-text"
Case "HJ文库"
wenkuwikiURL = "HJ文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as:off&ie=UTF8&keywords=%EF%BC%A8%EF%BC%AA%E6%96%87%E5%BA%AB&lo=digital-text&qid=1444268288&rh=n:2250738051,n:2275256051,n:2450063051,n:2410280051,k:%EF%BC%A8%EF%BC%AA%E6%96%87%E5%BA%AB,p_lbr_publishers_browse-bin:%E3%83%9B%E3%83%93%E3%83%BC%E3%82%B8%E3%83%A3%E3%83%91%E3%83%B3&sort=date-desc-rank"
Case "overlap文库"
wenkuwikiURL = "オーバーラップ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E3%82%AA%E3%83%BC%E3%83%90%E3%83%BC%E3%83%A9%E3%83%83%E3%83%97%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%82%AA%E3%83%BC%E3%83%90%E3%83%BC%E3%83%A9%E3%83%83%E3%83%97%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E3%82%AA%E3%83%BC%E3%83%90%E3%83%BC%E3%83%A9%E3%83%83%E3%83%97&qid=1448685097&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"
Case "幻冬舍金红石文库"
wenkuwikiURL = "幻冬舎ルチル文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2293148051%2Ck%3A%E5%B9%BB%E5%86%AC%E8%88%8E%E3%83%AB%E3%83%81%E3%83%AB%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E5%B9%BB%E5%86%AC%E8%88%8E%E3%82%B3%E3%83%9F%E3%83%83%E3%82%AF%E3%82%B9&sort=date-desc-rank&keywords=%E5%B9%BB%E5%86%AC%E8%88%8E%E3%83%AB%E3%83%81%E3%83%AB%E6%96%87%E5%BA%AB&ie=UTF8&qid=1449228689&lo=digital-text"
Case "角川Ruby文库"
wenkuwikiURL = "角川ルビー文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_nr_p_lbr_publishers_bro_0?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2293148051%2Ck%3A%E8%A7%92%E5%B7%9D%E3%83%AB%E3%83%93%E3%83%BC%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E8%A7%92%E5%B7%9D%E6%9B%B8%E5%BA%97&sort=date-desc-rank&keywords=%E8%A7%92%E5%B7%9D%E3%83%AB%E3%83%93%E3%83%BC%E6%96%87%E5%BA%AB&ie=UTF8&qid=1449228746&rnid=2256276051&lo=digital-text"
Case "B-PRINCE文库"
wenkuwikiURL = "B-PRINCE文庫"
newbookurl = "http://www.amazon.co.jp/gp/search/ref=sr_il_ti_digital-text?fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2293148051%2Ck%3AB-PRINCE%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3AKADOKAWA%2F%E3%82%A2%E3%82%B9%E3%82%AD%E3%83%BC%E3%83%BB%E3%83%A1%E3%83%87%E3%82%A3%E3%82%A2%E3%83%AF%E3%83%BC%E3%82%AF%E3%82%B9&sort=date-desc-rank&keywords=B-PRINCE%E6%96%87%E5%BA%AB&ie=UTF8&qid=1449157502&lo=digital-text"
Case "花丸文库"
wenkuwikiURL = "花丸文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E8%8A%B1%E4%B8%B8%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2293148051%2Ck%3A%E8%8A%B1%E4%B8%B8%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E7%99%BD%E6%B3%89%E7%A4%BE&qid=1449157794&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"
Case "白金文库"
wenkuwikiURL = "プラチナ文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E3%83%97%E3%83%A9%E3%83%81%E3%83%8A%E6%96%87%E5%BA%AB&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2293148051%2Ck%3A%E3%83%97%E3%83%A9%E3%83%81%E3%83%8A%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E3%83%97%E3%83%A9%E3%83%B3%E3%82%BF%E3%83%B3%E5%87%BA%E7%89%88&qid=1449157626&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"

Case "自由浪漫文库"
wenkuwikiURL = "フレジェロマンス文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E3%83%95%E3%83%AC%E3%82%B8%E3%82%A7%E3%83%AD%E3%83%9E%E3%83%B3%E3%82%B9%E6%96%87%E5%BA%AB&rh=n%3A2250738051%2Ck%3A%E3%83%95%E3%83%AC%E3%82%B8%E3%82%A7%E3%83%AD%E3%83%9E%E3%83%B3%E3%82%B9%E6%96%87%E5%BA%AB&qid=1449991083&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"

Case "甜爱文库"
wenkuwikiURL = "シュガーLOVE文庫"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_nr_p_lbr_publishers_bro_0?fst=as%3Aoff&rh=n%3A2250738051%2Ck%3A%E3%82%B7%E3%83%A5%E3%82%AC%E3%83%BCLOVE%E6%96%87%E5%BA%AB%2Cp_lbr_publishers_browse-bin%3A%E3%82%A4%E3%83%BC%E3%82%B9%E3%83%88%E3%83%BB%E3%83%97%E3%83%AC%E3%82%B9&sort=date-desc-rank&keywords=%E3%82%B7%E3%83%A5%E3%82%AC%E3%83%BCLOVE%E6%96%87%E5%BA%AB&ie=UTF8&qid=1449998832&rnid=2256276051&lo=digital-text"

Case "梅丽莎文库"
wenkuwikiURL = "メリッサ"
newbookurl = "http://www.amazon.co.jp/s/ref=sr_st_date-desc-rank?lo=digital-text&keywords=%E3%83%A1%E3%83%AA%E3%83%83%E3%82%B5&fst=as%3Aoff&rh=n%3A2250738051%2Cn%3A2275256051%2Cn%3A2450063051%2Cn%3A2410280051%2Ck%3A%E3%83%A1%E3%83%AA%E3%83%83%E3%82%B5%2Cp_lbr_publishers_browse-bin%3A%E4%B8%80%E8%BF%85%E7%A4%BE&qid=1449998936&__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&sort=date-desc-rank"

Case "list"
wenkuwikiURL = "分页"
newbookurl = ""
Case Else
wenkuwikiURL = sname
newbookurl = ""

End Select

End Function

Private Function getPub(publisher As String, Optional cstyle As String = "乙女向") As String
If cstyle = "乙女向" Then
Select Case publisher
Case "講談社"
    getPub = "讲谈社"
Case "KADOKAWA / 角川書店"
    getPub = "角川书店"
Case "KADOKAWA / アスキー?メディアワークス"
    getPub = "ASCII Media Works"
Case "KADOKAWA / アスキー·メディアワークス"
    getPub = "ASCII Media Works"
Case "KADOKAWA / エンターブレイン"
    getPub = "Enterbrain"
Case "イースト?プレス"
    getPub = "East Press"
Case "イースト·プレス"
    getPub = "East Press"
Case "コスミック出版"
    getPub = "Cosmic出版"
Case "ジュリアンパブリッシング"
    getPub = "Julian"
Case "ハーパーコリンズ?ジャパン"
    getPub = "Harper Collins"
Case "ハーパーコリンズ·ジャパン"
    getPub = "Harper Collins"
Case "プランタン出版"
    getPub = "Printemps出版"
Case "メディアソフト"
    getPub = "Media Soft"
Case "メディアックス"
    getPub = "Media Redox"
Case "オークラ出版"
    getPub = "Oakla出版"
Case "リブレ出版"
    getPub = "Libre出版"
Case "ハーレクイン"
    getPub = "Harlequin"
Case "幻冬舎" To "幻冬舎コミックス"
    getPub = "幻冬舍"
Case "ネットワーク出版"
    getPub = "Network出版"
Case "徳間書店（Chara）"
    getPub = "徳間書店"
Case "学研プラス"
    getPub = "学研+"
Case "スタジオプラスコ"
    getPub = "Studio Prasco"
Case "フロンティアワークス"
    getPub = "Frontier Works"
Case Else '白泉社 集英社
    getPub = publisher
End Select
ElseIf cstyle = "少年向" Then
Select Case publisher

Case "KADOKAWA / 角川書店"
    getPub = "角川書店"
Case "KADOKAWA / アスキー?メディアワークス"
    getPub = "ASCII Media Works"
Case "KADOKAWA / アスキー·メディアワークス"
    getPub = "ASCII Media Works"
Case "KADOKAWA / エンターブレイン"
    getPub = "Enter Brain"
Case "KADOKAWA / メディアファクトリー"
    getPub = "Media Factory"
Case "KADOKAWA / 富士見書房"
    getPub = "富士見書房"
Case "SBクリエイティブ"
    getPub = "SB Creative"
Case "講談社"
    getPub = "讲谈社"
Case "ホビージャパン"
    getPub = "Hobby Janpan"
Case "オーバーラップ"
    getPub = "overlap"
Case Else '小学館 一迅社  集英社
    getPub = publisher
End Select

End If
End Function
