/* global clearInterval, console, CustomFunctions, setInterval */
import { connectWS } from "../utils/websocket";

export class ETNET {
  static RTD(stockCode: string, type: "price" | "vol"): any {
    const ws = connectWS(stockCode, type);
    return {
      result: "等待数据...",
      stream: true
    };
  }
}

/**
 * Adds two numbers.
 * @customfunction
 * @param first First string
 * @param second Second string
 * @returns Data any.
 */
export function RTD2(stockCode: string, type: string,invocation: CustomFunctions.StreamingInvocation<string>): void {
  console.log("----RTD2 is called...")
  const ws = connectWS(stockCode, type,invocation);
  invocation.onCanceled = () => {
    
  };
}

export function clock2(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}


export function RTD2_BAK(stockCode: string, type: string): any {
  console.log("----RTD2 is called...")
 // const ws = connectWS(stockCode, type);
  return {
    result: "waiting...",
    stream: true
  };
}

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add2(first: number, second: number): number {
  return first + second + 100;
}

export function Time22(): any {
  return new Promise((resolve) => {
    // 初始值
    resolve(new Date().toLocaleTimeString());
    
    // 流式更新
    return (callback: (result: string) => void) => {
      const timer = setInterval(() => {
        callback(new Date().toLocaleTimeString());
      }, 1000); // 每秒更新

      // 清理函数
      return () => clearInterval(timer);
    };
  });
}  
/**
 * Adds two numbers.
 * @customfunction
 * @param first First string
 * @param second Second string
 * @returns 100.
 */
 
export function add3(stockCode: string, type: string): any {
 // const ws = connectWS(stockCode, type);
  return 100;
}

/**
 * Adds two numbers.
 * @customfunction
 * @returns 100.
 */
export function add4(): any {
  //const ws = connectWS(stockCode, type);
  return 100;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString('zh-CN', {
  hour12: false,
  hour: '2-digit',
  minute: '2-digit',
  second: '2-digit'
});
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

export function time(): string {
  return new Date().toLocaleTimeString();
}

export function rtd(stockCode: string, field: string): Promise<any> {
  return new Promise((resolve) => {
      const ws = new WebSocket("ws://localhost:8080/stock-data");
      ws.onmessage = (event) => {
          const data = JSON.parse(event.data);
          if (data.stockCode === stockCode && data.field === field) {
              resolve(data.value);
          }
      };
      ws.onopen = () => {
          ws.send(JSON.stringify({ stockCode, field }));
      };
  });
}

//1 注册自定义函数
CustomFunctions.associate("etnet.currentTime",currentTime);
CustomFunctions.associate("etnet.TIME",add2);
CustomFunctions.associate("etnet.RTD2",RTD2);
CustomFunctions.associate("etnet.RTD",rtd);
CustomFunctions.associate("etnet.add4",add4);
CustomFunctions.associate("etnet.Time22",add4);

