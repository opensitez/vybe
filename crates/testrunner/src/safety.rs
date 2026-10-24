//! Runner-side containment; guest execution and verdict semantics are unchanged.
use std::io;
use std::ops::{Deref, DerefMut};
use std::process::{Child as Process, Command};
use std::sync::{
    Arc, Condvar, Mutex, OnceLock,
    atomic::{AtomicU64, Ordering},
};
use std::thread::JoinHandle;
use std::time::Duration;

static MEMORY_LIMIT: OnceLock<u64> = OnceLock::new();
const RSS_SAMPLE_INTERVAL: Duration = Duration::from_millis(50);

#[derive(Default)]
struct WatchdogStop {
    requested: Mutex<bool>,
    wake: Condvar,
}

impl WatchdogStop {
    fn request(&self) {
        *self.requested.lock().unwrap() = true;
        self.wake.notify_all();
    }

    fn is_requested(&self) -> bool {
        *self.requested.lock().unwrap()
    }

    fn wait(&self, interval: Duration) -> bool {
        let requested = self.requested.lock().unwrap();
        self.wait_locked(requested, interval)
    }

    fn wait_locked(&self, requested: std::sync::MutexGuard<'_, bool>, interval: Duration) -> bool {
        // The predicate survives a request made before the monitor parks;
        // wait_timeout_while also tolerates spurious notifications.
        let (requested, _) = self.wake
            .wait_timeout_while(requested, interval, |requested| !*requested)
            .unwrap();
        *requested
    }
}

#[cfg(unix)]
static CHILD_GROUPS: [std::sync::atomic::AtomicI32; 1024] =
    [const { std::sync::atomic::AtomicI32::new(0) }; 1024];

#[cfg(unix)]
extern "C" fn interrupted(signal: libc::c_int) {
    // No locks, allocation, or formatting inside a signal handler.
    for group in &CHILD_GROUPS {
        let pid = group.load(Ordering::Relaxed);
        if pid > 1 {
            unsafe {
                libc::kill(-pid, libc::SIGKILL);
            }
        }
    }
    unsafe {
        libc::_exit(128 + signal);
    }
}

pub fn configure(memory_mb: u64) {
    let _ = MEMORY_LIMIT.set(memory_mb.saturating_mul(1024 * 1024));
    #[cfg(unix)]
    unsafe {
        libc::signal(libc::SIGINT, interrupted as *const () as libc::sighandler_t);
        libc::signal(
            libc::SIGTERM,
            interrupted as *const () as libc::sighandler_t,
        );
    }
}

pub struct Child {
    process: Process,
    stop: Arc<WatchdogStop>,
    exceeded: Arc<AtomicU64>,
    limit: u64,
    monitor: Option<JoinHandle<()>>,
    #[cfg(unix)]
    group_slot: usize,
}

impl Deref for Child {
    type Target = Process;
    fn deref(&self) -> &Process {
        &self.process
    }
}
impl DerefMut for Child {
    fn deref_mut(&mut self) -> &mut Process {
        &mut self.process
    }
}

impl Child {
    pub fn memory_failure(&self) -> Option<String> {
        let bytes = self.exceeded.load(Ordering::Relaxed);
        (bytes != 0).then(|| {
            format!(
                "worker memory limit exceeded: {} MiB (limit {} MiB); worker group killed",
                bytes / (1024 * 1024),
                self.limit / (1024 * 1024)
            )
        })
    }

    pub fn kill(&mut self) -> io::Result<()> {
        self.stop.request();
        kill_group(self.process.id());
        self.process.kill()
    }
}

impl Drop for Child {
    fn drop(&mut self) {
        let _ = self.kill();
        let _ = self.process.wait();
        if let Some(monitor) = self.monitor.take() {
            let _ = monitor.join();
        }
        #[cfg(unix)]
        CHILD_GROUPS[self.group_slot].store(0, Ordering::Relaxed);
    }
}

pub fn spawn(command: &mut Command) -> io::Result<Child> {
    spawn_limited(command, *MEMORY_LIMIT.get().unwrap_or(&(512 * 1024 * 1024)))
}

fn spawn_limited(command: &mut Command, limit: u64) -> io::Result<Child> {
    #[cfg(target_os = "macos")]
    {
        static CORE_LIMIT: OnceLock<Result<(), i32>> = OnceLock::new();
        // Inherited across posix_spawn. A pre_exec closure forces Rust's
        // fork/exec handshake, which can strand concurrent macOS worker boots.
        let result = CORE_LIMIT.get_or_init(|| {
            let core = libc::rlimit { rlim_cur: 0, rlim_max: 0 };
            if unsafe { libc::setrlimit(libc::RLIMIT_CORE, &core) } == 0 {
                Ok(())
            } else {
                Err(io::Error::last_os_error().raw_os_error().unwrap_or(libc::EIO))
            }
        });
        if let Err(errno) = result { return Err(io::Error::from_raw_os_error(*errno)); }
    }
    #[cfg(unix)]
    {
        use std::os::unix::process::CommandExt;
        command.process_group(0);
        // Only async-signal-safe syscalls in the post-fork child. The data
        // limit complements (not replaces) the resident-memory watchdog.
        #[cfg(not(target_os = "macos"))]
        unsafe {
            command.pre_exec(move || {
                let core = libc::rlimit {
                    rlim_cur: 0,
                    rlim_max: 0,
                };
                if libc::setrlimit(libc::RLIMIT_CORE, &core) != 0 {
                    return Err(io::Error::last_os_error());
                }
                #[cfg(target_os = "linux")]
                if limit != 0 {
                    let mut data: libc::rlimit = std::mem::zeroed();
                    if libc::getrlimit(libc::RLIMIT_DATA, &mut data) != 0 {
                        return Err(io::Error::last_os_error());
                    }
                    data.rlim_cur = data.rlim_cur.min(limit as libc::rlim_t);
                    if libc::setrlimit(libc::RLIMIT_DATA, &data) != 0 {
                        return Err(io::Error::last_os_error());
                    }
                }
                Ok(())
            });
        }
    }
    let mut process = command.spawn()?;
    let pid = process.id();
    #[cfg(unix)]
    let group_slot = match CHILD_GROUPS.iter().position(|slot| {
        slot.compare_exchange(0, pid as i32, Ordering::Relaxed, Ordering::Relaxed)
            .is_ok()
    }) {
        Some(slot) => slot,
        None => {
            kill_group(pid);
            let _ = process.kill();
            let _ = process.wait();
            return Err(io::Error::other("too many active test processes"));
        }
    };
    let stop = Arc::new(WatchdogStop::default());
    let exceeded = Arc::new(AtomicU64::new(0));
    let monitor = if limit != 0 {
        let stop = stop.clone();
        let exceeded = exceeded.clone();
        Some(std::thread::spawn(move || {
            while !stop.is_requested() {
                if let Some(bytes) = resident_bytes(pid) {
                    if bytes > limit {
                        exceeded.store(bytes, Ordering::Relaxed);
                        kill_group(pid);
                        break;
                    }
                }
                if stop.wait(RSS_SAMPLE_INTERVAL) {
                    break;
                }
            }
        }))
    } else {
        None
    };
    Ok(Child {
        process,
        stop,
        exceeded,
        limit,
        monitor,
        #[cfg(unix)]
        group_slot,
    })
}

fn kill_group(pid: u32) {
    #[cfg(unix)]
    if let Ok(pid) = i32::try_from(pid) {
        if pid > 1 {
            unsafe {
                libc::kill(-pid, libc::SIGKILL);
            }
        }
    }
    #[cfg(not(unix))]
    let _ = pid;
}

#[cfg(target_os = "macos")]
fn resident_bytes(pid: u32) -> Option<u64> {
    let mut usage: libc::rusage_info_v2 = unsafe { std::mem::zeroed() };
    let result = unsafe {
        libc::proc_pid_rusage(
            pid as i32,
            libc::RUSAGE_INFO_V2,
            (&mut usage as *mut libc::rusage_info_v2).cast(),
        )
    };
    (result == 0).then_some(usage.ri_resident_size.max(usage.ri_phys_footprint))
}

#[cfg(target_os = "linux")]
fn resident_bytes(pid: u32) -> Option<u64> {
    let status = std::fs::read_to_string(format!("/proc/{pid}/status")).ok()?;
    status.lines().find_map(|line| {
        line.strip_prefix("VmRSS:")?
            .split_whitespace()
            .next()?
            .parse::<u64>()
            .ok()?
            .checked_mul(1024)
    })
}

#[cfg(not(any(target_os = "linux", target_os = "macos")))]
fn resident_bytes(_: u32) -> Option<u64> {
    None
}

#[cfg(all(test, unix))]
mod tests {
    use super::*;

    #[test]
    fn watchdog_shutdown_before_wait_is_not_lost() {
        let stop = WatchdogStop::default();
        stop.request();
        assert!(stop.is_requested());
        assert!(stop.wait(Duration::ZERO));
        // Multiple kill/drop paths may request the same shutdown.
        stop.request();
        assert!(stop.wait(Duration::ZERO));
    }

    #[test]
    fn watchdog_shutdown_interrupts_long_wait_and_acknowledges() {
        let stop = Arc::new(WatchdogStop::default());
        let (ready_tx, ready_rx) = std::sync::mpsc::channel();
        let (ack_tx, ack_rx) = std::sync::mpsc::channel();
        let monitor_stop = stop.clone();
        let monitor = std::thread::spawn(move || {
            // Holding the predicate lock through readiness makes request()
            // acquire it only after wait_locked atomically parks/releases it.
            let requested = monitor_stop.requested.lock().unwrap();
            ready_tx.send(()).unwrap();
            let stopped = monitor_stop.wait_locked(requested, Duration::from_secs(60));
            ack_tx.send(stopped).unwrap();
        });
        ready_rx.recv_timeout(Duration::from_secs(5)).unwrap();
        stop.request();
        // This generous failure deadline is not a sub-50ms speed assertion:
        // the acknowledgment proves notification interrupted a 60-second wait.
        let acknowledgment = ack_rx.recv_timeout(Duration::from_secs(5));
        assert!(matches!(acknowledgment, Ok(true)));
        monitor.join().unwrap();
    }

    #[test]
    fn watchdog_sample_timeout_does_not_request_shutdown() {
        let stop = WatchdogStop::default();
        assert!(!stop.wait(Duration::ZERO));
        assert!(!stop.is_requested());
        assert_eq!(RSS_SAMPLE_INTERVAL, Duration::from_millis(50));
        stop.request();
        assert!(stop.wait(RSS_SAMPLE_INTERVAL));
    }

    #[test]
    fn worker_has_its_own_group_and_drop_reaps_it() {
        let mut command = Command::new("/bin/sleep");
        command.arg("10");
        let child = spawn_limited(&mut command, 128 * 1024 * 1024).unwrap();
        let pid = child.id() as i32;
        let slot = child.group_slot;
        assert_eq!(unsafe { libc::getpgid(pid) }, pid);
        assert_ne!(unsafe { libc::getpgrp() }, pid);
        drop(child);
        assert_eq!(unsafe { libc::kill(pid, 0) }, -1);
        assert_ne!(CHILD_GROUPS[slot].load(Ordering::Relaxed), pid);
        let mut status = 0;
        assert_eq!(unsafe { libc::waitpid(pid, &mut status, libc::WNOHANG) }, -1);
        assert_eq!(io::Error::last_os_error().raw_os_error(), Some(libc::ECHILD));
    }

    #[test]
    fn disabled_watchdog_still_reaps_and_unregisters_child() {
        let mut command = Command::new("/bin/sleep");
        command.arg("10");
        let child = spawn_limited(&mut command, 0).unwrap();
        let pid = child.id() as i32;
        let slot = child.group_slot;
        assert!(child.monitor.is_none());
        assert!(child.memory_failure().is_none());
        drop(child);
        assert_eq!(unsafe { libc::kill(pid, 0) }, -1);
        assert_ne!(CHILD_GROUPS[slot].load(Ordering::Relaxed), pid);
    }

    #[test]
    fn unix_child_disables_core_dumps() {
        use std::io::Read;
        let mut command = Command::new("/bin/sh");
        command
            .args(["-c", "ulimit -c"])
            .stdout(std::process::Stdio::piped());
        let mut child = spawn_limited(&mut command, 128 * 1024 * 1024).unwrap();
        let mut output = String::new();
        child
            .stdout
            .take()
            .unwrap()
            .read_to_string(&mut output)
            .unwrap();
        assert!(child.wait().unwrap().success());
        assert_eq!(output.trim(), "0");
    }

    #[cfg(target_os = "macos")]
    #[test]
    fn rss_watchdog_kills_a_small_child_without_a_large_allocation() {
        use wait_timeout::ChildExt;
        let mut command = Command::new("/bin/sleep");
        command.arg("10");
        // A one-byte budget tests the watchdog using ordinary executable RSS;
        // it does not deliberately exhaust memory to prove containment.
        let mut child = spawn_limited(&mut command, 1).unwrap();
        let status = child.wait_timeout(Duration::from_secs(2)).unwrap();
        if status.is_none() {
            let _ = child.kill();
        }
        assert!(status.is_some(), "RSS monitor failed to stop child");
        assert!(child.memory_failure().is_some());
    }
}
