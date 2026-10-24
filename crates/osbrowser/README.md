# osbrowser

[webcore](https://crates.io/crates/webcore)'s API, with the system browser doing the rendering.

`osbrowser` exposes the same modules, types and signatures as webcore:
`types::Document`, `dom`, `window`, `browser`, `canvas`, `ui_events`,
`scheduling` and `platform::dialogs`. A program written against one compiles
against the other with a one-word change. With webcore, the page is laid out
and painted by webcore, in your process. With osbrowser, it's shown in the
user's own browser (Chrome, Firefox, Safari, or whatever the platform opens for
a URL), and your program drives it over a loopback channel.

No browser engine, UI toolkit or rasteriser is linked. The page runs where the
user already has a browser.

```toml
[dependencies]
osbrowser = "0.5"
```

```rust
use osbrowser::browser::{Browser, BrowserDocument, BrowserGuest, BrowserOptions};
use osbrowser::dom::events::ListenerOptions;

struct Page;
impl BrowserGuest for Page {
    fn dispatch_pending(&mut self) -> bool { false }
}

fn main() {
    let document = BrowserDocument::new(osbrowser::load_html("<html><body></body></html>", 800.0));
    {
        let mut doc = document.write();
        let body = doc.body().unwrap();
        let button = doc.create_element("button");
        doc.set_text_content(button, "Click me");
        doc.append_child(body, button);
        doc.add_event_listener(button, "click", Box::new(|_event, doc| {
            doc.set_title("clicked");
        }), ListenerOptions::default());
    }
    let mut browser = Browser::open(BrowserOptions::default()).unwrap();
    browser.attach_document(document);
    browser.set_guest(Page);
    browser.run().unwrap();
}
```

Run the full example with `cargo run --example counter`.

## How it works

- **The DOM lives in the browser.** `Document` is this process's handle on
  it. Node ids follow webcore's conventions: a `u32`, with `0` meaning "no node".
- **Writes don't wait.** Mutations go into one ordered queue and aren't waited
  on. A read is a single round trip that also delivers every write queued
  before it. New nodes get their ids in this process, so creating one costs no
  round trip.
- **Listeners run here.** `add_event_listener` keeps the closure in this
  process. The browser reports each event it delivers, and the handler runs in
  `Browser::run`'s loop, or before a dispatching call such as
  `process_mouse_event` returns.
- **The canvas API is webcanvas's**, the same crate webcore re-exports, so
  the types are identical. Each `<canvas>` keeps a `RecordingCanvas` here: state
  reads (`get_transform`, `fill_style_css`, `line_width`, …) never leave the
  process, and the commands one paint records go to the page's real `<canvas>`
  as a single message. Only `measure_text`, hit testing, `get_image_data` and
  `to_data_url` ask the browser.

## Differences from webcore

- `renderer::Renderer` has no in-process frame to paint, because the browser
  composites into its own window. `render` leaves the pixmap untouched and says
  so on stderr.
- A listener can't cancel the browser's own default action for a real user
  event. By the time the event reaches this process, the browser has already
  acted on it.

## License

GPL-3.0-or-later, or a commercial license, the same as webcore.
