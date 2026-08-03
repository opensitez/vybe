//! Project file format (vybe.toml) and multi-language compilation.
//!
//! A vybe.toml project defines:
//! ```toml
//! [project]
//! name = "MyApp"
//! entry = "main.vb"
//!
//! [sources]
//! files = ["main.vb", "utils.vb", "engine.js"]
//!
//! [window]
//! title = "My App"
//! width = 800
//! height = 600
//!
//! [host]
//! gui = true
//! filesystem = true
//! ```

#[derive(Debug, Clone)]
pub struct ProjectConfig {
    pub name: String,
    pub entry: String,
    pub files: Vec<String>,
    pub window: Option<WindowConfig>,
    pub host: HostConfig }

#[derive(Debug, Clone)]
pub struct WindowConfig {
    pub title: String,
    pub width: u32,
    pub height: u32 }

#[derive(Debug, Clone, Default)]
pub struct HostConfig {
    pub gui: bool,
    pub filesystem: bool,
    pub network: bool,
    pub database: bool }

impl ProjectConfig {
    /// Parse a simple TOML-like project file.
    pub fn parse(content: &str) -> Result<Self, String> {
        let mut name = String::from("VybeProject");
        let mut entry = String::new();
        let mut files: Vec<String> = Vec::new();
        let mut window = None;
        let mut host = HostConfig::default();

        let mut section = "";
        let mut win_title = String::new();
        let mut win_width = 800u32;
        let mut win_height = 600u32;

        for line in content.lines() {
            let line = line.trim();
            if line.is_empty() || line.starts_with('#') {
                continue;
            }

            if line.starts_with('[') && line.ends_with(']') {
                section = match &line[1..line.len() - 1] {
                    "project" => "project",
                    "sources" => "sources",
                    "window" => "window",
                    "host" => "host",
                    _ => "" };
                continue;
            }

            if let Some(eq) = line.find('=') {
                let key = line[..eq].trim().to_lowercase();
                let val = line[eq + 1..].trim().trim_matches('"');

                match section {
                    "project" => match key.as_str() {
                        "name" => name = val.to_string(),
                        "entry" => entry = val.to_string(),
                        _ => {}
                    },
                    "sources" => {
                        if key == "files" {
                            // Parse array: ["a.vb", "b.js"]
                            let inner = val.trim_start_matches('[').trim_end_matches(']');
                            for item in inner.split(',') {
                                let f = item.trim().trim_matches('"').trim_matches('\'');
                                if !f.is_empty() {
                                    files.push(f.to_string());
                                }
                            }
                        }
                    }
                    "window" => match key.as_str() {
                        "title" => win_title = val.to_string(),
                        "width" => win_width = val.parse().unwrap_or(800),
                        "height" => win_height = val.parse().unwrap_or(600),
                        _ => {}
                    },
                    "host" => match key.as_str() {
                        "gui" => host.gui = val == "true",
                        "filesystem" => host.filesystem = val == "true",
                        "network" => host.network = val == "true",
                        "database" => host.database = val == "true",
                        _ => {}
                    },
                    _ => {}
                }
            }
        }

        if !win_title.is_empty() {
            window = Some(WindowConfig {
                title: win_title,
                width: win_width,
                height: win_height });
        }

        if entry.is_empty() && !files.is_empty() {
            entry = files[0].clone();
        }

        Ok(ProjectConfig {
            name,
            entry,
            files,
            window,
            host })
    }
}
