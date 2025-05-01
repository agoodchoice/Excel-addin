const sockets = new Map<string, WebSocket>();

export function connectWS(stockCode: string, type: string,invocation: CustomFunctions.StreamingInvocation<string>): WebSocket {
  const key = `${stockCode}-${type}`;
  console.log("-------connectWS is called....step 01----");
  if (!sockets.has(key)) {
    // const ws = new WebSocket("ws://localhost:2049/"); //wss
    const ws = new WebSocket("ws://localhost:2050/"); //wss
    //2050
    // console.log("-------WebSocket is create....-step 02---");
    ws.onmessage = (event) => {
      console.log("-------onmessage is -..step 033----, event.data=" + event.data);
      const data1: string ='5';
      const [header, data] = event.data.split('\n');
      const datas = event.data.split('\n');
      //let fields = data.split(',');
      let fields = datas[datas.length -1].split(',');
      const value_last = fields[fields.length -2];

      console.log("-------onmessage is -..step 033----, fields=" + fields +", value_last=" + value_last);
      // data1 = '5';
      //const context = Excel.CustomFunctionHelpers.context;
      //const context = Office.context;
      //const range = context.workbook.getSelectedRange();
     // const context = Excel.RequestContext.workbook.getSelectedRange();
     // const range = context.workbook.getSelectedRange();
     invocation.setResult("4.15");
     invocation.setResult(value_last);
    };    

    ws.onopen = () => {
      console.log("-------onmessage is called....step 02----");
      // ws.send("-1,20,0,0,ETNET.RTDSERVER,User,1.0.0.0");
      // console.log("-------onmessage is called....step 05----");
      //ws.send("-1,1,0,2,");
      ws.send("-1,1,0,2," +stockCode +"," + type );
      //ws.send("-1,1,0,2,700,tc");
      // console.log("-------onmessage is called....step 06----");
    };
    //ws.onopen = () => ws.send("-1,1,0,2,3,cd");
    sockets.set(key, ws);
  }
  return sockets.get(key)!;
}