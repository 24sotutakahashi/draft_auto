function myFunction() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getActiveSheet();　//アクティブなスプレッドシート・シートを取得

  let row = sheet.getLastRow()-1;　//アクティブな行数を取得、列の行はいらないので-1してあげる
  let column = sheet.getLastColumn();　//アクティブな列数を取得、

  let range = sheet.getRange(2,1,row,column); //2行1列目から、セルに値がある範囲までのデータを指定
  let data = range.getValues(); //値を取得
  //var html_data = HtmlService.createTemplateFromFile("index").evaluate().getContent();

  for(let i = 0; i < data.length; i++){

    if(data[i][7] == ""){
    continue;
    }
    let recipient = data[i][7]; // 宛先
    //console.log(recipient);

    let subject = "ご協賛のお願い" // 件名
    let body = " "; // 本文

    //let MailText =  "abcde";
    let NAME = data[i][6];

    let url_video = "https://www.youtube.com/watch?v=tMgIk70cw1s";
    let url_annual = "https://www.thoughtfulgift.org/_files/ugd/0daab0_0a4ae142f2974fb59e2b28472436b76c.pdf";
    let url_flyer = "https://www.thoughtfulgift.org/_files/ugd/0daab0_fe8465c4c83e43eb94b24b9086e258ca.pdf";
    let url_Tokyo = "https://www.tokyo-np.co.jp/article/185910"

    // 本文
    let html =`${NAME}<br />
    お問い合わせご担当者様<br />
    &nbsp;<br />
    初めてご連絡させていただきます。<br />
    私たちは2021年4月より精神障害により精神科入院する方へ入院中に必要な物を無償で提供する非営利活動をしております。<br />
    神奈川県内13病院に入院する60名以上の方をすでに支援いたしました。<br />
    各方面からの注目もあり、神奈川新聞社会面や<a href="${url_Tokyo}">東京新聞</a>取り上げていただいたり<br />
    パルシステム神奈川で行った賛助金カンパでは172人の方から賛助金をいただきました。<br />
    詳しい活動内容はWebサイトの<a href="${url_video}">動画</a>や<a href="${url_annual}">アニュアルレポート</a>をご参考いただけあすと幸いです。<br />
    &nbsp;<br />
    御社で商品在庫をSDGｓ活動に活かせないか考えている、<br />
    過去の大会のTシャツやノベルティグッズが余っている等ございませんでしょうか？<br />
    ございましたらぜひ私たちの活動にご協力をお願いします。&nbsp;<br />
    (企業/団体様からのご協賛をお願いする<a href="${url_flyer}">チラシ</a>)<br />
    ご協賛を受けるにあたって弁護士が作成した合意書等のご準備もございます。<br />
    &nbsp;<br />
    実際に今受けているご寄付は、引越用段ボールの提供、TVCM制作会社から衣装の提供<br />
    神奈川県/横浜市社会福祉協議会からのタオル・歯ブラシ等がございます。<br />
    &nbsp;<br />
    ご検討いただけますと幸いです。<br />
    どうぞよろしくお願いいたします。<br />
    &nbsp;<br />
一般社団法人Thoughtful&nbsp;Gift<br />
　((自分の名前))<br />
　TEL＆FAX&nbsp;：044(385)3107<br />
　URL：https://www.thoughtfulgift.org/<br />
&nbsp;&nbsp;&nbsp;&nbsp;E-mail：thoughtfulgift.cp@gmail.com`;

    GmailApp.createDraft(recipient, subject, body, {htmlBody: html});
    };
    
    
    }
