use crate::project::Project;
use std::fs;
use std::path::Path;

#[derive(Debug, thiserror::Error)]
pub enum ProjectError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),
}

pub type ProjectResult<T> = Result<T, ProjectError>;

pub fn save_project(project: &Project, path: impl AsRef<Path>) -> ProjectResult<()> {
    let json = serde_json::to_string_pretty(project)?;
    fs::write(path, json)?;
    Ok(())
}

pub fn load_project(path: impl AsRef<Path>) -> ProjectResult<Project> {
    let json = fs::read_to_string(path)?;
    let project = serde_json::from_str(&json)?;
    Ok(project)
}
