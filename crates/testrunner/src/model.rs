//! Result model — testecma's (`ecma/testecma/src/lib.rs`) generalised across
//! languages.
//!
//! Same five outcomes and the same aggregate/report shape, because the point of
//! that shape is run-over-run comparison and it already works. What changes is
//! the grouping axis: test262 aggregates by *feature* (declared in each test's
//! YAML frontmatter); we have no frontmatter, so we aggregate by the two things
//! a path already tells us — language and category.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Hash)]
#[serde(rename_all = "lowercase")]
pub enum TestResult {
    Pass,
    Fail,
    Skip,
    Timeout,
    Error,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TestExecution {
    pub path: PathBuf,
    pub language: String,
    pub category: String,
    pub name: String,
    pub result: TestResult,
    pub message: String,
    pub duration_ms: u128,
}

impl TestExecution {
    /// `go/json_marshal/marshal_bool_true` — the identity used in reports and
    /// the one a person types to re-run a single case.
    pub fn slug(&self) -> String {
        format!("{}/{}/{}", self.language, self.category, self.name)
    }
}

/// Split `tests/<lang>/<category>/<name>.<ext>` into its parts.
pub fn identify(path: &Path) -> (String, String, String) {
    let name = path
        .file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_default();
    let category = path
        .parent()
        .and_then(|p| p.file_name())
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_default();
    let language = path
        .parent()
        .and_then(|p| p.parent())
        .and_then(|p| p.file_name())
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_default();
    (language, category, name)
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GroupStats {
    pub total: usize,
    pub passed: usize,
    pub failed: usize,
    pub skipped: usize,
    pub errors: usize,
}

impl GroupStats {
    pub fn pass_rate(&self) -> f64 {
        if self.total == 0 {
            0.0
        } else {
            (self.passed as f64 / self.total as f64) * 100.0
        }
    }

    fn record(&mut self, result: TestResult) {
        self.total += 1;
        match result {
            TestResult::Pass => self.passed += 1,
            TestResult::Fail | TestResult::Timeout => self.failed += 1,
            TestResult::Skip => self.skipped += 1,
            TestResult::Error => self.errors += 1,
        }
    }
}

#[derive(Debug, Serialize, Deserialize)]
pub struct TestReport {
    pub timestamp: String,
    /// What the tests ran under — `target/debug/vybex`, `go run`, `python3`.
    /// A report is only comparable to another with the same runtime.
    pub runtime: String,
    pub total: usize,
    pub passed: usize,
    pub failed: usize,
    pub skipped: usize,
    pub errors: usize,
    pub timeouts: usize,
    pub by_result: HashMap<String, usize>,
    pub by_language: HashMap<String, GroupStats>,
    pub by_category: HashMap<String, GroupStats>,
    pub executions: Vec<TestExecution>,
    pub duration_secs: u64,
}

impl TestReport {
    pub fn new(runtime: String) -> Self {
        Self {
            timestamp: chrono::Local::now().to_rfc3339(),
            runtime,
            total: 0,
            passed: 0,
            failed: 0,
            skipped: 0,
            errors: 0,
            timeouts: 0,
            by_result: Default::default(),
            by_language: Default::default(),
            by_category: Default::default(),
            executions: Vec::new(),
            duration_secs: 0,
        }
    }

    pub fn add_execution(&mut self, exec: TestExecution) {
        self.total += 1;
        match exec.result {
            TestResult::Pass => self.passed += 1,
            TestResult::Fail => self.failed += 1,
            TestResult::Timeout => {
                self.failed += 1;
                self.timeouts += 1;
            }
            TestResult::Skip => self.skipped += 1,
            TestResult::Error => self.errors += 1,
        }

        let key = format!("{:?}", exec.result).to_lowercase();
        *self.by_result.entry(key).or_insert(0) += 1;

        self.by_language
            .entry(exec.language.clone())
            .or_default()
            .record(exec.result);
        self.by_category
            .entry(format!("{}/{}", exec.language, exec.category))
            .or_default()
            .record(exec.result);

        self.executions.push(exec);
    }

    pub fn pass_rate(&self) -> f64 {
        if self.total == 0 {
            0.0
        } else {
            (self.passed as f64 / self.total as f64) * 100.0
        }
    }

    /// Failing slugs, sorted — the set a comparison diffs to name what newly
    /// broke and what newly passes, rather than only moving a count.
    pub fn failing_slugs(&self) -> Vec<String> {
        let mut slugs: Vec<String> = self
            .executions
            .iter()
            .filter(|e| matches!(e.result, TestResult::Fail | TestResult::Timeout | TestResult::Error))
            .map(|e| e.slug())
            .collect();
        slugs.sort();
        slugs
    }
}
