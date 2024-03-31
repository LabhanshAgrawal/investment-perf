// ==UserScript==
// @name         Zerodha Order History
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  try to take over the world!
// @author       You
// @match        https://coin.zerodha.com/dashboard/mf/orders
// @icon         https://www.google.com/s2/favicons?sz=64&domain=zerodha.com
// @grant        none
// ==/UserScript==

const f = async function () {
  let from = "2019-01-01";
  const today = new Date().toISOString().split("T")[0];
  const result = [];

  while (from <= today) {
    const to = new Date(new Date(from).getTime() + 1000 * 60 * 60 * 24 * 30)
      .toISOString()
      .split("T")[0];

    const data = await getData(from, to <= today ? to : today);
    console.log({from,to,data});
    result.push(...data);

    from = new Date(new Date(from).getTime() + 1000 * 60 * 60 * 24 * 31)
      .toISOString()
      .split("T")[0];

    await new Promise(resolve => setTimeout(resolve, 5000));
  }

  const csvString = [
    Object.keys(result[0]),
    ...result.map(item => Object.values(item)),
  ]
    .map(e => e.join(","))
    .join("\n");

  console.log(csvString);
};

async function getData(from, to) {
  const res = await fetch(
    `https://coin.zerodha.com/api/mf/orders?from=${from}&to=${to}`,
    {
      headers: {
        accept: "application/json, text/plain, */*",
        "accept-language": "en-US,en;q=0.9",
        "sec-ch-ua":
          '"Chromium";v="104", " Not A;Brand";v="99", "Google Chrome";v="104"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"macOS"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "x-csrftoken": document.cookie
          .split(";")
          .map(v => v.split("="))
          .reduce((acc, v_1) => {
            acc[decodeURIComponent(v_1[0].trim())] = decodeURIComponent(
              v_1[1].trim()
            );
            return acc;
          }, {}).public_token,
      },
      referrer: "https://coin.zerodha.com/dashboard/mf/orders",
      referrerPolicy: "strict-origin-when-cross-origin",
      body: null,
      method: "GET",
      mode: "cors",
      credentials: "include",
    }
  );
  const data = await res.json();
  return data.data;
}


const d = document.createElement("div");
d.style.position = "absolute";
d.style.top = "0px";
d.style.left = "0px";
d.style.width = '30px';
d.style.height = '30px';
d.style.background = 'red';
d.style.zIndex = '1000';
d.style.cursor = 'pointer';
d.addEventListener('click', () => {
  d.remove();
  f();
}
);
document.body.appendChild(d);

