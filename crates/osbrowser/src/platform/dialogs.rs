//! User-agent dialogs, as webcore's `platform::dialogs` exposes them.
//!
//! `alert` and `confirm` are the page's own, so they appear over the page the
//! user is looking at. The file pickers are the operating system's, which is
//! what a browser shows for them too.

use std::path::PathBuf;

use serde_json::json;

use crate::page;

pub fn alert(message: &str) {
    page::read::<bool>(page::visible_page(), "window_alert", json!([message]));
}

pub fn confirm(message: &str) -> bool {
    page::read(page::visible_page(), "window_confirm", json!([message]))
}

fn dialog(title: &str, filters: &[(String, Vec<String>)], directory: &str) -> rfd::FileDialog {
    let mut dialog = rfd::FileDialog::new().set_title(title);
    if !directory.is_empty() {
        dialog = dialog.set_directory(directory);
    }
    for (name, extensions) in filters {
        let extensions: Vec<&str> = extensions.iter().map(String::as_str).collect();
        dialog = dialog.add_filter(name, &extensions);
    }
    dialog
}

pub fn open_file(
    title: &str,
    filters: &[(String, Vec<String>)],
    directory: &str,
    multiple: bool,
) -> Vec<PathBuf> {
    let dialog = dialog(title, filters, directory);
    if multiple {
        dialog.pick_files().unwrap_or_default()
    } else {
        dialog.pick_file().into_iter().collect()
    }
}

pub fn save_file(
    title: &str,
    filters: &[(String, Vec<String>)],
    directory: &str,
    suggested: &str,
) -> Option<PathBuf> {
    let mut dialog = dialog(title, filters, directory);
    if !suggested.is_empty() {
        dialog = dialog.set_file_name(suggested);
    }
    dialog.save_file()
}

pub fn pick_directory(title: &str, directory: &str) -> Option<PathBuf> {
    dialog(title, &[], directory).pick_folder()
}
