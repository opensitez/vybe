//! The browser window and its event loop, as webcore's `browser` exposes it.
//!
//! The window is a tab in the system browser. [`Browser::run`] is the loop
//! webcore's window runs: deliver what the page reported to the listeners,
//! let the guest run what is due, and sleep until the next event or deadline.

mod document;

pub use document::{BrowserDocument, DocumentRead, DocumentWrite};

use std::time::Duration;

use crate::dom::registry;
use crate::page;

pub trait BrowserGuest {
    fn context_ready(&mut self) {}
    fn context_closed(&mut self) {}
    fn dispatch_pending(&mut self) -> bool;
    fn next_wakeup(&self) -> Option<Duration> {
        None
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct BrowserChrome {
    pub tabs: bool,
    pub address_bar: bool,
    pub navigation_buttons: bool,
}

impl Default for BrowserChrome {
    fn default() -> Self {
        Self {
            tabs: true,
            address_bar: true,
            navigation_buttons: true,
        }
    }
}

impl BrowserChrome {
    pub const NONE: Self = Self {
        tabs: false,
        address_bar: false,
        navigation_buttons: false,
    };
}

#[derive(Clone, Debug, Default)]
pub struct BrowserDebugOptions {
    pub port: Option<u16>,
    pub chrome_port: u16,
    pub chrome_srgb: bool,
}

#[derive(Clone, Debug)]
pub struct BrowserOptions {
    pub title: String,
    pub width: f32,
    pub height: f32,
    pub chrome: BrowserChrome,
    pub default_context_menu: bool,
    pub debug: BrowserDebugOptions,
}

impl Default for BrowserOptions {
    fn default() -> Self {
        Self {
            title: "osbrowser".into(),
            width: 1280.0,
            height: 900.0,
            chrome: BrowserChrome::default(),
            default_context_menu: true,
            debug: BrowserDebugOptions::default(),
        }
    }
}

/// Why the browser could not be opened or run.
#[derive(Clone, Debug)]
pub struct BrowserError(String);

impl std::fmt::Display for BrowserError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for BrowserError {}

/// How long the loop sleeps when nothing is due: one frame.
const FRAME: Duration = Duration::from_millis(16);

pub struct Browser {
    options: BrowserOptions,
    guest: Option<Box<dyn BrowserGuest>>,
}

impl Browser {
    pub fn open(options: BrowserOptions) -> Result<Self, BrowserError> {
        page::session().map_err(BrowserError)?;
        Ok(Self { options, guest: None })
    }

    pub fn attach_document(&mut self, document: BrowserDocument) -> crate::window::WindowId {
        let id = registry::register_document(document);
        crate::window::adopt(id, &self.options.title)
    }

    pub fn set_guest(&mut self, guest: impl BrowserGuest + 'static) {
        self.guest = Some(Box::new(guest));
    }

    pub fn run(mut self) -> Result<(), BrowserError> {
        let session = page::session().map_err(BrowserError)?;
        let mut guest = self.guest.take();
        session.begin_batch();
        page::end_startup_batch();
        if let Some(guest) = guest.as_mut() {
            guest.context_ready();
        }
        session.end_batch();
        while page::connected() {
            session.begin_batch();
            for document in registry::open_documents() {
                document.write_trusted().pump();
            }
            if let Some(guest) = guest.as_mut() {
                guest.dispatch_pending();
            }
            session.end_batch();
            let wait = guest
                .as_ref()
                .and_then(|guest| guest.next_wakeup())
                .map_or(FRAME, |due| due.min(FRAME));
            session.wait_for_event(wait);
        }
        if let Some(guest) = guest.as_mut() {
            guest.context_closed();
        }
        Ok(())
    }
}

/// Declare that this process will not present its documents in a window.
///
/// Must precede the first document: its page is then hosted by a headless
/// browser process, with no tab and no window.
pub fn set_headless(headless: bool) {
    page::set_headless(headless);
}
