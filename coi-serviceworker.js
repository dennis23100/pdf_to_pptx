/*! coi-serviceworker v0.1.7 - Modified for simplicity */
let coep = 'require-corp';
let coop = 'same-origin';

if (typeof window === 'undefined') {
  self.addEventListener("install", () => self.skipWaiting());
  self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()));

  self.addEventListener("fetch", function (event) {
    if (event.request.cache === "only-if-cached" && event.request.mode !== "same-origin") {
      return;
    }

    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response.status === 0) {
            return response;
          }

          const newHeaders = new Headers(response.headers);
          newHeaders.set("Cross-Origin-Embedder-Policy", coep);
          newHeaders.set("Cross-Origin-Opener-Policy", coop);

          return new Response(response.body, {
            status: response.status,
            statusText: response.statusText,
            headers: newHeaders,
          });
        })
        .catch((e) => console.error(e))
    );
  });
} else {
  (() => {
    // 檢查是否已經開啟隔離模式，是的話就不用做事了
    if (window.crossOriginIsolated) {
        console.log('Site is cross-origin isolated (COOP+COEP).');
        return;
    }

    const n = navigator;
    if (n.serviceWorker) {
        n.serviceWorker.register(window.document.currentScript.src).then(
            (registration) => {
                console.log("COI Service Worker registered");
                
                // 如果 Service Worker 剛啟動，重新整理頁面以讓設定生效
                registration.addEventListener("updatefound", () => {
                    window.location.reload();
                });

                if (registration.active && !n.serviceWorker.controller) {
                    window.location.reload();
                }
            },
            (err) => {
                console.error("COI Service Worker failed to register", err);
            }
        );
    }
  })();
}
