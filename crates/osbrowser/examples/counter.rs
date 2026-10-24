//! A page built through webcore's API and shown in the system browser.
//!
//! Replace `osbrowser` with `webcore` and the same program runs in webcore's
//! own window instead — that is the point of the crate.
//!
//!     cargo run --example counter

use osbrowser::browser::{Browser, BrowserDocument, BrowserGuest, BrowserOptions};
use osbrowser::dom::events::ListenerOptions;

/// Nothing of its own to run: the page's listeners do the work.
struct Page;

impl BrowserGuest for Page {
    fn dispatch_pending(&mut self) -> bool {
        false
    }
}

fn main() {
    let document = BrowserDocument::new(osbrowser::load_html(
        "<html><head><title>Counter</title></head>\
         <body style='font: 16px sans-serif; margin: 24px'></body></html>",
        800.0,
    ));
    {
        let mut doc = document.write();
        let body = doc.body().expect("a parsed page has a body");
        let button = doc.create_element("button");
        doc.set_text_content(button, "Click me");
        doc.append_child(body, button);
        let label = doc.create_element("p");
        doc.set_text_content(label, "0 clicks");
        doc.append_child(body, label);

        let mut clicks = 0;
        doc.add_event_listener(
            button,
            "click",
            Box::new(move |_event, doc| {
                clicks += 1;
                doc.set_text_content(label, &format!("{clicks} clicks"));
            }),
            ListenerOptions::default(),
        );
    }

    let mut browser = Browser::open(BrowserOptions { title: "Counter".into(), ..Default::default() })
        .expect("open the browser");
    browser.attach_document(document);
    browser.set_guest(Page);
    browser.run().expect("run the browser");
}
