use std::time::{Duration, Instant};

#[derive(Default)]
pub(super) struct RunTimer {
    resumed: Option<Instant>,
}

impl RunTimer {
    pub(super) fn resume(&mut self, now: Instant) {
        self.resumed.get_or_insert(now);
    }

    pub(super) fn stop(&mut self, now: Instant) -> Option<Duration> {
        self.resumed.take().map(|start| now.saturating_duration_since(start))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn intervals_exclude_paused_inspection_time() {
        let start = Instant::now();
        let mut timer = RunTimer::default();
        assert_eq!(timer.stop(start), None);
        timer.resume(start);
        assert_eq!(timer.stop(start + Duration::from_millis(250)),
            Some(Duration::from_millis(250)));
        let resumed = start + Duration::from_secs(60);
        timer.resume(resumed);
        assert_eq!(timer.stop(resumed + Duration::from_millis(500)),
            Some(Duration::from_millis(500)));
        assert_eq!(timer.stop(resumed + Duration::from_secs(1)), None);
    }

    #[test]
    fn duplicate_resume_does_not_restart_an_interval() {
        let start = Instant::now();
        let mut timer = RunTimer::default();
        timer.resume(start);
        timer.resume(start + Duration::from_secs(1));
        assert_eq!(timer.stop(start + Duration::from_secs(2)),
            Some(Duration::from_secs(2)));
    }
}
