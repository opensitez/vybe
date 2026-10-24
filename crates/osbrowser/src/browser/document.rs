//! Shared ownership of a live DOM between a browser context and its host bridge.

use std::ops::{Deref, DerefMut};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, MutexGuard};

use crate::Document;

#[derive(Clone)]
pub struct BrowserDocument {
    document: Arc<Mutex<Document>>,
    resources_dirty: Arc<AtomicBool>,
    external_writes_pending: Arc<AtomicBool>,
    wake: Arc<Mutex<Option<Arc<dyn Fn() + Send + Sync>>>>,
}

impl BrowserDocument {
    pub fn new(document: Document) -> Self {
        Self {
            document: Arc::new(Mutex::new(document)),
            resources_dirty: Arc::new(AtomicBool::new(true)),
            external_writes_pending: Arc::new(AtomicBool::new(false)),
            wake: Arc::new(Mutex::new(None)),
        }
    }

    pub fn read(&self) -> DocumentRead<'_> {
        DocumentRead::Shared(
            self.document
                .lock()
                .expect("browser document lock poisoned"),
        )
    }

    pub fn write(&self) -> DocumentWrite<'_> {
        DocumentWrite::ExternalShared {
            document: Some(self.document.lock().expect("browser document lock poisoned")),
            changed: false,
            pending: &self.external_writes_pending,
            wake: &self.wake,
        }
    }

    pub(crate) fn write_trusted(&self) -> DocumentWrite<'_> {
        DocumentWrite::Shared(
            self.document
                .lock()
                .expect("browser document lock poisoned"),
        )
    }

    pub fn resources_changed(&self) {
        self.resources_dirty.store(true, Ordering::Release);
    }

    /// Wake the browser after an external producer queues document work.
    pub fn notify_pending_tasks(&self) {
        let wake = self.wake.lock().expect("browser wake lock poisoned").clone();
        if let Some(wake) = wake {
            wake();
        }
    }

    pub(crate) fn same_document(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.document, &other.document)
    }
}

pub enum DocumentRead<'a> {
    Borrowed(&'a Document),
    Shared(MutexGuard<'a, Document>),
}

impl Deref for DocumentRead<'_> {
    type Target = Document;
    fn deref(&self) -> &Document {
        match self {
            Self::Borrowed(document) => document,
            Self::Shared(document) => document,
        }
    }
}

pub enum DocumentWrite<'a> {
    Borrowed(&'a mut Document),
    ExternalBorrowed {
        document: Option<&'a mut Document>,
        changed: bool,
        pending: &'a mut bool,
        wake: Option<&'a Arc<dyn Fn() + Send + Sync>>,
    },
    Shared(MutexGuard<'a, Document>),
    ExternalShared {
        document: Option<MutexGuard<'a, Document>>,
        changed: bool,
        pending: &'a AtomicBool,
        wake: &'a Mutex<Option<Arc<dyn Fn() + Send + Sync>>>,
    },
}

impl Deref for DocumentWrite<'_> {
    type Target = Document;
    fn deref(&self) -> &Document {
        match self {
            Self::Borrowed(document) => document,
            Self::Shared(document) => document,
            Self::ExternalBorrowed { document, .. } => document.as_deref().expect("live document guard"),
            Self::ExternalShared { document, .. } => document.as_deref().expect("live document guard"),
        }
    }
}

impl DerefMut for DocumentWrite<'_> {
    fn deref_mut(&mut self) -> &mut Document {
        match self {
            Self::Borrowed(document) => document,
            Self::Shared(document) => document,
            Self::ExternalBorrowed { document, changed, .. } => {
                *changed = true;
                document.as_deref_mut().expect("live document guard")
            }
            Self::ExternalShared { document, changed, .. } => {
                *changed = true;
                document.as_deref_mut().expect("live document guard")
            }
        }
    }
}

impl Drop for DocumentWrite<'_> {
    fn drop(&mut self) {
        if let Self::ExternalBorrowed { document, changed, pending, wake } = self {
            drop(document.take());
            if *changed && !**pending {
                **pending = true;
                if let Some(callback) = wake { callback(); }
            }
        }
        if let Self::ExternalShared { document, changed, pending, wake } = self {
            drop(document.take());
            if *changed && !pending.swap(true, Ordering::AcqRel) {
                let callback = wake.lock().expect("browser wake lock poisoned").clone();
                if let Some(callback) = callback { callback(); }
            }
        }
    }
}

