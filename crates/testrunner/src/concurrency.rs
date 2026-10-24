//! Worker-pool sizing is separate from per-process resource enforcement.

pub const DEFAULT_JOBS: usize = 12;

pub fn default_jobs() -> usize {
    DEFAULT_JOBS
}

#[cfg(test)]
mod tests {
    #[test]
    fn default_pool_has_twelve_active_workers() {
        assert_eq!(super::default_jobs(), 12);
    }
}
