//! Native OS dialogs — message boxes, file pickers, folder pickers.
//!
//! Thin wrapper over the `rfd` crate that exposes a clean Rust API in
//! `vybe_widgets`'s own namespace, so consumers (the CLI form runner,
//! the host bridge, framework wrappers) don't need to depend on `rfd`
//! directly. Concentrating the dependency here also means we can swap
//! the backend later (xdg-portal, AppKit, win32 directly, …) without
//! touching every caller.
//!
//! All functions are blocking. The OS handles the modal loop; the
//! caller's thread waits until the user dismisses the dialog. Async
//! variants are intentionally NOT exposed — every existing caller is
//! synchronous (the CLI form runner blocks the VM tick on a host call,
//! which is itself blocking).

use std::path::PathBuf;

// ─── Message boxes ─────────────────────────────────────────────────────────

/// Severity of a message box dialog.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MessageLevel {
    Info,
    Warning,
    Error,
}

/// User's choice from a yes/no/ok-cancel dialog.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DialogChoice {
    Ok,
    Cancel,
    Yes,
    No,
    /// User dismissed the dialog (X button, Esc, etc.) without
    /// choosing.
    Dismissed,
}

/// Native message box dialog.
///
/// Use one of the convenience constructors (`info`, `warning`, `error`)
/// for the common single-button-OK case, or build a tailored dialog via
/// the `new` + `with_*` builder methods for yes/no / ok-cancel.
pub struct MessageBox {
    text: String,
    title: String,
    level: MessageLevel,
    buttons: MessageButtons,
}

#[derive(Clone, Copy)]
enum MessageButtons {
    Ok,
    OkCancel,
    YesNo,
}

impl MessageBox {
    /// Show an informational message box with an OK button. Blocks
    /// until the user dismisses.
    pub fn info(text: impl Into<String>, title: impl Into<String>) {
        Self::new(text, title).with_level(MessageLevel::Info).show();
    }

    /// Show a warning message box with an OK button.
    pub fn warning(text: impl Into<String>, title: impl Into<String>) {
        Self::new(text, title)
            .with_level(MessageLevel::Warning)
            .show();
    }

    /// Show an error message box with an OK button.
    pub fn error(text: impl Into<String>, title: impl Into<String>) {
        Self::new(text, title)
            .with_level(MessageLevel::Error)
            .show();
    }

    /// Show a yes/no question. Returns `Yes`, `No`, or `Dismissed`.
    pub fn yes_no(text: impl Into<String>, title: impl Into<String>) -> DialogChoice {
        Self::new(text, title)
            .with_level(MessageLevel::Info)
            .with_yes_no()
            .show_with_choice()
    }

    /// Show an OK / Cancel question. Returns `Ok`, `Cancel`, or
    /// `Dismissed`.
    pub fn ok_cancel(text: impl Into<String>, title: impl Into<String>) -> DialogChoice {
        Self::new(text, title)
            .with_level(MessageLevel::Info)
            .with_ok_cancel()
            .show_with_choice()
    }

    /// Build a custom message box. The default level is `Info` and
    /// the default button set is `Ok`.
    pub fn new(text: impl Into<String>, title: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            title: title.into(),
            level: MessageLevel::Info,
            buttons: MessageButtons::Ok,
        }
    }

    pub fn with_level(mut self, level: MessageLevel) -> Self {
        self.level = level;
        self
    }

    pub fn with_yes_no(mut self) -> Self {
        self.buttons = MessageButtons::YesNo;
        self
    }

    pub fn with_ok_cancel(mut self) -> Self {
        self.buttons = MessageButtons::OkCancel;
        self
    }

    /// Show the dialog and ignore the result. Use [`show_with_choice`]
    /// if you care about which button was pressed.
    ///
    /// [`show_with_choice`]: Self::show_with_choice
    pub fn show(self) {
        let _ = self.show_with_choice();
    }

    /// Show the dialog and return the user's choice.
    pub fn show_with_choice(self) -> DialogChoice {
        let level = match self.level {
            MessageLevel::Info => rfd::MessageLevel::Info,
            MessageLevel::Warning => rfd::MessageLevel::Warning,
            MessageLevel::Error => rfd::MessageLevel::Error,
        };
        let buttons = match self.buttons {
            MessageButtons::Ok => rfd::MessageButtons::Ok,
            MessageButtons::OkCancel => rfd::MessageButtons::OkCancel,
            MessageButtons::YesNo => rfd::MessageButtons::YesNo,
        };
        let result = rfd::MessageDialog::new()
            .set_title(&self.title)
            .set_description(&self.text)
            .set_level(level)
            .set_buttons(buttons)
            .show();
        match result {
            rfd::MessageDialogResult::Yes => DialogChoice::Yes,
            rfd::MessageDialogResult::No => DialogChoice::No,
            rfd::MessageDialogResult::Ok => DialogChoice::Ok,
            rfd::MessageDialogResult::Cancel => DialogChoice::Cancel,
            _ => DialogChoice::Dismissed,
        }
    }
}

// ─── File / folder dialogs ─────────────────────────────────────────────────

/// One filter for a file dialog: a human-readable name plus a list of
/// extensions (without leading dots).
///
/// Example: `FileFilter::new("Images", &["png", "jpg", "jpeg"])`.
#[derive(Clone, Debug)]
pub struct FileFilter {
    pub name: String,
    pub extensions: Vec<String>,
}

impl FileFilter {
    pub fn new(name: impl Into<String>, extensions: &[&str]) -> Self {
        Self {
            name: name.into(),
            extensions: extensions.iter().map(|s| s.to_string()).collect(),
        }
    }
}

/// Native file picker — open or save.
///
/// ```ignore
/// use vybe_widgets::dialogs::{FileDialog, FileFilter};
///
/// if let Some(path) = FileDialog::new("Open Image")
///     .with_filter(FileFilter::new("Images", &["png", "jpg"]))
///     .open()
/// {
///     // load `path`
/// }
/// ```
pub struct FileDialog {
    title: String,
    filters: Vec<FileFilter>,
    starting_dir: Option<PathBuf>,
    starting_file: Option<String>,
}

impl FileDialog {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            filters: Vec::new(),
            starting_dir: None,
            starting_file: None,
        }
    }

    pub fn with_filter(mut self, filter: FileFilter) -> Self {
        self.filters.push(filter);
        self
    }

    pub fn with_filters(mut self, filters: impl IntoIterator<Item = FileFilter>) -> Self {
        self.filters.extend(filters);
        self
    }

    pub fn with_starting_directory(mut self, dir: impl Into<PathBuf>) -> Self {
        self.starting_dir = Some(dir.into());
        self
    }

    pub fn with_filename(mut self, name: impl Into<String>) -> Self {
        self.starting_file = Some(name.into());
        self
    }

    /// Show an "open file" dialog. Blocks until the user picks a file
    /// or cancels.
    pub fn open(self) -> Option<PathBuf> {
        self.build().pick_file()
    }

    /// Show an "open multiple files" dialog.
    pub fn open_multiple(self) -> Option<Vec<PathBuf>> {
        self.build().pick_files()
    }

    /// Show a "save file" dialog.
    pub fn save(self) -> Option<PathBuf> {
        self.build().save_file()
    }

    fn build(self) -> rfd::FileDialog {
        let mut d = rfd::FileDialog::new().set_title(&self.title);
        for f in &self.filters {
            let exts: Vec<&str> = f.extensions.iter().map(|s| s.as_str()).collect();
            d = d.add_filter(&f.name, &exts);
        }
        if let Some(dir) = self.starting_dir.as_ref() {
            d = d.set_directory(dir);
        }
        if let Some(name) = self.starting_file.as_ref() {
            d = d.set_file_name(name);
        }
        d
    }
}

/// Native folder picker.
///
/// ```ignore
/// use vybe_widgets::dialogs::FolderDialog;
///
/// if let Some(dir) = FolderDialog::new("Select Project Folder").pick() {
///     // load `dir`
/// }
/// ```
pub struct FolderDialog {
    title: String,
    starting_dir: Option<PathBuf>,
}

impl FolderDialog {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            starting_dir: None,
        }
    }

    pub fn with_starting_directory(mut self, dir: impl Into<PathBuf>) -> Self {
        self.starting_dir = Some(dir.into());
        self
    }

    /// Show the folder picker and return the chosen folder, or `None`
    /// if the user cancelled.
    pub fn pick(self) -> Option<PathBuf> {
        let mut d = rfd::FileDialog::new().set_title(&self.title);
        if let Some(dir) = self.starting_dir.as_ref() {
            d = d.set_directory(dir);
        }
        d.pick_folder()
    }

    /// Show a multi-select folder picker.
    pub fn pick_multiple(self) -> Option<Vec<PathBuf>> {
        let mut d = rfd::FileDialog::new().set_title(&self.title);
        if let Some(dir) = self.starting_dir.as_ref() {
            d = d.set_directory(dir);
        }
        d.pick_folders()
    }
}
