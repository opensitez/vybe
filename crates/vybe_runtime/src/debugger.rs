//! In-VM step debugger — the execution-side half of the debug surface described
//! in `fulldebugplan.md`. This module owns the pause/step/breakpoint state
//! machine and the typed request/response/event protocol. The transport (TCP /
//! REPL / IDE) lives in `vybex` and only ever moves these types across channels.
//!
//! Design invariants (see the plan):
//!   * The VM dispatch loop calls [`Debugger::on_instruction`] at most once per
//!     instruction, and only when `self.instrumented` is set — zero cost off.
//!   * Pausing BLOCKS the dispatch thread on `cmd_rx.recv()` (true suspend, not a
//!     busy loop). `continue`/`step`/`detach` unblock it.
//!   * Inspection is read-only against a coherent, whole-instruction-boundary
//!     snapshot of VM state.
//!   * No opcode, no execution-semantics change: the debugger observes and gates.

use std::sync::mpsc::{Receiver, Sender};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use std::cell::RefCell;
use std::collections::{HashMap, VecDeque};
use std::sync::atomic::{AtomicU32, AtomicU64, AtomicUsize, Ordering};

use crate::{VMError, Value};
use crate::opcode::Op;
use crate::vm::VM;

/// Timing and current phase shared with the debugger client. Compiler work can
/// block the VM thread, so a report must read this without a VM command.
#[derive(Default)]
pub struct DebugReport {
    pub process_started: Option<Instant>,
    pub milestones: Vec<(String, Duration)>,
    pub first_output: Option<Duration>,
    pub last_output: Option<Duration>,
    pub paused: Duration,
    pub active: Vec<(String, Instant)>,
    pub completed: Vec<(String, Duration)>,
    pub active_host: Option<(Arc<str>, Instant)>,
    pub active_host_args: Option<Vec<String>>,
    pub location: Option<String>,
    pub instructions: u64,
    /// One function sample per 4096 executed VM instructions.
    pub function_samples: HashMap<String, u64>,
    /// One completed host call in 64, measured at the existing debugger
    /// boundary. Durations are reported separately from VM instruction samples.
    pub host_samples: HashMap<Arc<str>, (u64, Duration)>,
    pub host_sample_tick: u64,
    /// Source-line samples expose hot statements without an external profiler.
    pub line_samples: HashMap<String, u64>,
    pub live: Arc<DebugLiveInstruction>,
}

#[derive(Default)]
pub struct DebugLiveInstruction {
    pub chunk: AtomicUsize,
    pub ip: AtomicUsize,
    pub op: AtomicU32,
    pub instructions: AtomicU64,
}

pub type SharedDebugReport = Arc<Mutex<DebugReport>>;

impl DebugReport {
    pub fn milestone(&mut self, name: &str) {
        if let Some(start) = self.process_started {
            self.milestones.push((name.to_string(), start.elapsed()));
        }
    }
}

/// Called by the stdout host after a nonempty write; includes execute inside
/// the same debugger report scope as their caller.
pub fn record_stdout_write() {
    CURRENT_REPORT.with(|slot| {
        if let Some(report) = slot.borrow().as_ref() {
            let mut report = report.lock().unwrap();
            if let Some(start) = report.process_started {
                let elapsed = start.elapsed();
                report.first_output.get_or_insert(elapsed);
                report.last_output = Some(elapsed);
            }
        }
    });
}

struct DebugPauseTime(SharedDebugReport, Instant);
impl Drop for DebugPauseTime {
    fn drop(&mut self) {
        self.0.lock().unwrap().paused += self.1.elapsed();
    }
}

thread_local! {
    static CURRENT_REPORT: RefCell<Option<SharedDebugReport>> = const { RefCell::new(None) };
}

/// Makes compiler phases visible to the attached debugger on this VM thread.
pub struct DebugReportScope(Option<SharedDebugReport>);

impl DebugReportScope {
    pub fn enter(report: Option<SharedDebugReport>) -> Self {
        let previous = CURRENT_REPORT.with(|slot| slot.replace(report));
        Self(previous)
    }
}

impl Drop for DebugReportScope {
    fn drop(&mut self) {
        CURRENT_REPORT.with(|slot| { slot.replace(self.0.take()); });
    }
}

pub struct DebugPhase {
    report: Option<SharedDebugReport>,
    label: String,
    start: Instant,
}

/// Shows a blocking host call to `report` without retaining a timing entry
/// for every ordinary host call in a long-running program.
pub struct DebugHostCall {
    report: Option<SharedDebugReport>,
    previous: Option<(Arc<str>, Instant)>,
    previous_args: Option<Vec<String>>,
}

impl DebugHostCall {
    pub fn new(report: Option<SharedDebugReport>, name: impl Into<Arc<str>>) -> Self {
        let (previous, previous_args) = if let Some(report) = &report {
            let mut report = report.lock().unwrap();
            let previous = report.active_host.replace((name.into(), Instant::now()));
            (previous, report.active_host_args.take())
        } else {
            (None, None)
        };
        Self { report, previous, previous_args }
    }

    /// Snapshot only scalar arguments before the host call starts. Object
    /// formatting could take locks or traverse arbitrarily large guest data.
    pub fn args(&self, args: &[Value]) {
        if let Some(report) = &self.report {
            let preview = args.iter().take(8).map(|value| match value {
                Value::I32(v) => v.to_string(),
                Value::I64(v) => v.to_string(),
                Value::F64(v) => v.to_string(),
                Value::String(v) => format!("string(len={})", v.len()),
                Value::Null => "null".to_string(),
                Value::Undefined => "undefined".to_string(),
                Value::Object(obj) => format!("object({:p})", Arc::as_ptr(obj)),
                other => format!("{:?}", std::mem::discriminant(other)),
            }).collect();
            report.lock().unwrap().active_host_args = Some(preview);
        }
    }

    pub fn stage(&self, name: &str) {
        if let Some(report) = &self.report {
            report.lock().unwrap().active_host = Some((Arc::from(name), Instant::now()));
        }
    }
}

impl Drop for DebugHostCall {
    fn drop(&mut self) {
        if let Some(report) = &self.report {
            let mut report = report.lock().unwrap();
            report.host_sample_tick += 1;
            if report.host_sample_tick % 64 == 0 {
                if let Some((name, elapsed)) = report.active_host.as_ref()
                    .map(|(name, start)| (name.clone(), start.elapsed())) {
                    let sample = report.host_samples.entry(name).or_default();
                    sample.0 += 1;
                    sample.1 += elapsed;
                }
            }
            report.active_host = self.previous.take();
            report.active_host_args = self.previous_args.take();
        }
    }
}

impl DebugPhase {
    pub fn current(label: impl Into<String>) -> Self {
        let report = CURRENT_REPORT.with(|slot| slot.borrow().clone());
        Self::new(report, label)
    }
    pub fn current_lazy(label: impl FnOnce() -> String) -> Self {
        let report = CURRENT_REPORT.with(|slot| slot.borrow().clone());
        let label = if report.is_some() { label() } else { String::new() };
        Self::new(report, label)
    }
    pub fn new(report: Option<SharedDebugReport>, label: impl Into<String>) -> Self {
        let label = label.into();
        let start = Instant::now();
        if let Some(report) = &report {
            report.lock().unwrap().active.push((label.clone(), start));
        }
        Self { report, label, start }
    }
}

impl Drop for DebugPhase {
    fn drop(&mut self) {
        if let Some(report) = &self.report {
            let mut report = report.lock().unwrap();
            if let Some(index) = report.active.iter().rposition(|(label, _)| label == &self.label) {
                report.active.remove(index);
            }
            report.completed.push((self.label.clone(), self.start.elapsed()));
        }
    }
}

// ─── Protocol: client → VM ──────────────────────────────────────────────────

/// A command from the client, paired with a one-shot channel for its reply.
pub struct DebugRequest {
    pub command: DebugCommand,
    pub reply: Sender<DebugResponse>,
}

/// Where a breakpoint / step should land.
#[derive(Debug, Clone)]
pub enum ChunkRef {
    Index(usize),
    Name(String),
}

#[derive(Debug, Clone)]
pub enum DebugCommand {
    // ── Execution control ──
    Continue,
    /// Pause at the next instruction boundary (async "interrupt").
    Pause,
    /// Step to the next instruction, descending into calls.
    StepInto,
    /// Step over calls (pause when back at ≤ current frame depth).
    StepOver,
    /// Run until the current frame returns (pause at strictly shallower depth).
    StepOut,
    /// Single bytecode instruction (same as StepInto here; distinct name for UIs).
    StepInstruction,
    /// Detach the debugger and let the program run free.
    Detach,
    /// Terminate the program.
    Quit,

    // ── Breakpoints ──
    /// Break at a source line within a chunk. Optional `condition` fires the
    /// breakpoint only when the expression is truthy in the paused frame.
    BreakLine {
        chunk: ChunkRef,
        line: u32,
        condition: Option<String>,
    },
    /// Break at a bytecode offset within a chunk.
    BreakOffset {
        chunk: ChunkRef,
        offset: usize,
        condition: Option<String>,
    },
    /// Break at a source line across ALL chunks (file:line, file-agnostic). Sets
    /// one breakpoint per chunk that has an instruction on that line.
    BreakSourceLine {
        line: u32,
        condition: Option<String>,
    },
    /// Break on entry to the function named `name`.
    BreakFunction {
        name: String,
        condition: Option<String>,
    },
    /// Break before a call to the named imported host function.
    BreakHost { name: Option<String>, bad_type_only: bool },
    /// Logpoint: log a message (with `{expr}` interpolation) at a source line and
    /// keep running — never pauses.
    Logpoint {
        line: u32,
        message: String,
    },
    /// Run to a source line once, then remove the breakpoint (run-to-cursor).
    RunToLine {
        line: u32,
    },
    /// Skip a breakpoint's first `count` hits.
    SetIgnoreCount {
        id: u32,
        count: u32,
    },
    /// Break when an exception is thrown (`on_throw`) or only when it would be
    /// uncaught (`on_uncaught`). Both false disables.
    ExceptionBreak {
        on_throw: bool,
        on_uncaught: bool,
    },
    ListBreakpoints,
    /// Remove every breakpoint (DAP `setBreakpoints` replace semantics).
    ClearBreakpoints,
    DeleteBreakpoint {
        id: u32,
    },
    EnableBreakpoint {
        id: u32,
        enabled: bool,
    },

    // ── Inspection (valid while paused) ──
    /// Read a variable by name in the current frame (local via debug names,
    /// else global), with optional `.field` / `[index]` structural drill-down.
    Print {
        path: String,
    },
    /// Write a literal (number / string / bool / null) into a named local or
    /// global in the current frame.
    SetVar {
        name: String,
        literal: String,
    },
    Backtrace,
    Locals {
        frame: usize,
    },
    OperandStack,
    Globals {
        prefix: Option<String>,
    },
    /// Report module global initializer decisions and resolved values.
    TraceGlobalInits { enabled: bool },
    /// Pause before a named global initializer (`*` matches all names).
    BreakGlobalInit { name: String },
    /// Disassemble around the current ip or an explicit offset in this chunk.
    Disasm {
        window: usize,
        offset: Option<usize>,
    },
    /// Show structured-control branches whose destination is near an offset.
    ControlTargets { offset: usize },
    /// List chunks (index, name, arity).
    Chunks,
    /// The FRAME'S SHAPE: every slot the frame actually has, compiler-owned
    /// ones included, next to the chunk metadata that decides the layout.
    ///
    /// `locals` deliberately shows only user names and hides `__*` as
    /// "compiler scratch" — which hides exactly the slots that break. A
    /// capturing function's closure environment is `__capture_0`, so when it
    /// was being overwritten by emitter scratch, `locals` reported
    /// "(no locals)" for a frame with 48 live slots. This command answers the
    /// question that one cannot: what is really in this frame, and does the
    /// layout the VM filled match the layout the compiler assigned.
    Frame {
        frame: usize,
    },

    /// Stateful hot reload: recompile the source and swap changed function
    /// bodies in place, preserving heap/globals/stack (Dart-style, stage 1).
    Reload,
    /// Restart the whole program from scratch (fresh state).
    Restart,

    // ── Data watchpoints ──
    AddWatchpoint {
        target: String,
    },
    ListWatchpoints,
    ClearWatchpoints,

    // ── Fibers / threads ──
    /// List the current + suspended fibers (threads analog).
    Fibers,

    // ── Watch expressions (re-evaluated on every pause) ──
    AddWatch {
        expr: String,
    },
    ListWatches,
    ClearWatches,

    // ── Live event stream (VYBE_TRACE replacement) ──
    /// Turn per-instruction `opcode` events on/off.
    StreamOpcodes {
        enabled: bool,
    },
    /// Toggle whether the entry pause skips the runtime prelude (default on).
    SetSkipSystem {
        enabled: bool,
    },
    /// Simulate a GUI event: invoke `control`'s `event` handler through the live
    /// VM (e.g. a button Click or window Close) without an OS window.
    FireEvent {
        control: String,
        event: String,
    },
}

// ─── Protocol: VM → client (replies) ────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum DebugResponse {
    Ok,
    Error(String),
    /// A rendered value (from `Print`) or a before→after note (from `SetVar`).
    Value(String),
    BreakpointSet {
        id: u32,
        chunk: String,
        offset: usize,
        line: Option<u32>,
    },
    Breakpoints(Vec<BreakpointInfo>),
    Backtrace(Vec<FrameInfo>),
    Locals(Vec<SlotInfo>),
    OperandStack(Vec<String>),
    Globals(Vec<(String, String)>),
    Disasm {
        current_ip: usize,
        lines: Vec<DisasmLine>,
    },
    Chunks(Vec<ChunkInfo>),
    Frame(FrameShape),
    /// One summary line per live fiber (current + suspended).
    Fibers(Vec<String>),
}

// ─── Protocol: VM → client (async events) ───────────────────────────────────

#[derive(Debug, Clone)]
pub enum DebugEvent {
    /// Execution stopped; the client can now inspect and control.
    Paused {
        reason: PauseReason,
        location: Location,
        frame_summary: String,
        /// Current values of watch expressions (`expr`, rendered-or-error).
        watches: Vec<(String, String)>,
    },
    /// Execution resumed after a pause.
    Resumed,
    /// The program finished. `value` is the rendered result.
    Exited { value: String },
    /// A per-instruction trace event (only when opcode streaming is on).
    Opcode {
        chunk: String,
        ip: usize,
        op: String,
        stack_depth: usize,
    },
    /// An informational message from the debugger.
    Log { message: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PauseReason {
    Entry,
    Breakpoint {
        id: u32,
    },
    Step,
    Interrupt,
    /// A watchpoint's value changed.
    Watchpoint {
        id: u32,
    },
    /// About to execute a throw (exception breakpoint). `uncaught` = no handler.
    Exception {
        uncaught: bool,
    },
    HostCall { name: String },
}

#[derive(Debug, Clone)]
pub struct Location {
    pub chunk_index: usize,
    pub chunk_name: String,
    pub ip: usize,
    pub line: Option<u32>,
}

#[derive(Debug, Clone)]
pub struct FrameInfo {
    pub depth: usize,
    pub chunk_index: usize,
    pub chunk_name: String,
    pub ip: usize,
    pub line: Option<u32>,
}

#[derive(Debug, Clone)]
pub struct SlotInfo {
    pub index: usize,
    /// Source variable name for this slot, if the compiler emitted debug names.
    pub name: Option<String>,
    pub value: String,
}

#[derive(Debug, Clone)]
pub struct DisasmLine {
    pub offset: usize,
    pub text: String,
    pub is_current: bool,
    pub line: Option<u32>,
}

#[derive(Debug, Clone)]
pub struct BreakpointInfo {
    pub id: u32,
    pub chunk_index: usize,
    pub chunk_name: String,
    pub offset: usize,
    pub line: Option<u32>,
    pub enabled: bool,
}

#[derive(Debug, Clone)]
pub struct ChunkInfo {
    pub index: usize,
    pub name: String,
    pub arity: u8,
    pub code_len: usize,
}

/// A frame's declared layout beside its actual contents.
///
/// The four counts are the compiler's answer to "how big is this frame and
/// what lives where"; the slots are what the VM actually holds. When those two
/// disagree the program misbehaves with no trap at the fault, so they are
/// shown together on purpose.
#[derive(Debug, Clone)]
pub struct FrameShape {
    pub frame: usize,
    pub chunk_index: usize,
    pub chunk_name: String,
    pub arity: u8,
    pub local_count: u16,
    pub scratch_high_water: u16,
    /// First slot the VM fills with upvalues at call time, and how many.
    /// `capture_count > 0` with `capture_base` inside the parameter range is a
    /// layout bug, not a quirk.
    pub capture_base: u16,
    pub capture_count: u8,
    pub base: usize,
    /// Every slot in `0..local_count`, named where a name exists. `owner`
    /// distinguishes a user binding from a compiler-owned slot so scratch is
    /// visible rather than filtered.
    pub slots: Vec<FrameSlot>,
}

#[derive(Debug, Clone)]
pub struct FrameSlot {
    pub index: usize,
    pub name: Option<String>,
    pub owner: SlotOwner,
    pub value: String,
    /// False when the slot lies beyond the frame's live stack extent — the
    /// lazy-locals convention means an untouched slot may not exist yet.
    pub live: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SlotOwner {
    /// A binding the program declared.
    User,
    /// The closure environment the VM fills at entry (`__capture_*`).
    Capture,
    /// Any other compiler-generated slot: scratch, `__shared_env`, temporaries.
    Compiler,
}

impl SlotOwner {
    pub fn tag(self) -> &'static str {
        match self {
            SlotOwner::User => "user",
            SlotOwner::Capture => "capture",
            SlotOwner::Compiler => "compiler",
        }
    }
}

// ─── Debugger state ─────────────────────────────────────────────────────────

struct Breakpoint {
    id: u32,
    chunk_index: usize,
    /// Function-entry breakpoints match by name so they also cover chunks
    /// appended by eval, include, or hot reload after the command was set.
    function_name: Option<String>,
    offset: usize,
    enabled: bool,
    /// Optional condition expression — the breakpoint fires only when it
    /// evaluates truthy in the paused frame (compiler-faithful via `debug_eval`).
    condition: Option<String>,
    /// Skip this many hits before it starts pausing (`ignore <id> <n>`).
    ignore_count: u32,
    /// Running count of times the location was reached.
    hit_count: u32,
    /// Logpoint: when set, emit this message (with `{expr}` interpolation) and
    /// auto-continue instead of pausing.
    log_message: Option<String>,
    /// Remove after firing once (run-to-cursor).
    one_shot: bool,
}

/// A data/value watchpoint: pause when the value of `target` (a variable name or
/// `a.b[0]` path) changes. Checked once per instruction while armed.
struct Watchpoint {
    id: u32,
    target: String,
    last: String,
}

#[derive(Clone, Copy, PartialEq)]
enum RunMode {
    /// Running free; still services async requests (pause/set-bp) each instruction.
    Running,
    /// Stopped, blocked on the command channel.
    Paused,
    /// Pause at the very next instruction (StepInto / StepInstruction / interrupt).
    PauseNext,
    /// Pause when frame depth returns to ≤ target on the same fiber (StepOver).
    StepOver { target_depth: usize, fiber: u64 },
    /// Pause when frame depth drops below target on the same fiber (StepOut).
    StepOut { target_depth: usize, fiber: u64 },
}

/// The execution-side debugger. Held as `Option<Debugger>` on the VM and taken
/// out for the duration of each `on_instruction` call to avoid borrow conflicts.
pub struct Debugger {
    pub report: SharedDebugReport,
    host_names: HashMap<usize, Arc<str>>,
    live: Arc<DebugLiveInstruction>,
    sampled_instructions: u64,
    recent_instructions: VecDeque<(usize, usize, Op)>,
    cmd_rx: Receiver<DebugRequest>,
    evt_tx: Sender<DebugEvent>,
    breakpoints: Vec<Breakpoint>,
    next_bp_id: u32,
    watchpoints: Vec<Watchpoint>,
    next_wp_id: u32,
    break_on_throw: bool,
    break_on_uncaught: bool,
    break_host: Option<(String, bool)>,
    mode: RunMode,
    stream_opcodes: bool,
    trace_global_inits: bool,
    break_global_inits: Vec<String>,
    watches: Vec<String>,
    /// Set once the first instruction has been seen (so PauseNext doesn't fire
    /// on the same instruction it was requested from).
    armed: bool,
    /// When true (default), the entry pause is skipped over the injected runtime
    /// prelude and lands on the first line of the user's own code
    /// (`<script>.user_code_offset`). Toggle with `skip-system off` to debug the
    /// prelude/runtime. Only affects the initial entry pause; explicit
    /// breakpoints inside the prelude still fire.
    skip_system: bool,
    /// One-time latch so the prelude-skip only arms on the first instruction.
    prelude_skip_done: bool,
    /// A resuming command (`continue`/`step`/`next`/`out`/`run-to`) whose reply is
    /// held until the VM next pauses — see `Flow::Resume { await_stop }`. Flushed
    /// in `enter_pause`; if the program instead runs to completion the `Sender`
    /// drops with the debugger, unblocking the client via a closed channel.
    pending_reply: Option<Sender<DebugResponse>>,
    /// True during the prelude-skip auto-run (entry → first user-code pause).
    /// While set, `on_instruction` does not service client commands, so a piped
    /// command batch waits in the channel and is processed in order at the first
    /// real pause instead of being consumed mid-run. Cleared in `enter_pause`.
    defer_cmds_until_pause: bool,
}

impl Debugger {
    /// Host indices are stable within this VM. Resolve their display names
    /// once instead of scanning the entire plugin registry on every host call.
    pub(crate) fn host_name(&mut self, index: usize, registry: &HashMap<(String, String), usize>) -> Arc<str> {
        self.host_names.entry(index).or_insert_with(|| {
            registry.iter().find(|(_, id)| **id == index)
                .map(|((module, name), _)| Arc::from(format!("{module}:{name}")))
                .unwrap_or_else(|| Arc::from(format!("host #{index}")))
        }).clone()
    }
    /// Keep a finished CLI session available for an in-process restart. The
    /// reader thread closes the channel on piped EOF, so scripted debugger
    /// runs still exit promptly after their output.
    pub(crate) fn wait_after_exit(&mut self) -> bool {
        while let Ok(req) = self.cmd_rx.recv() {
            match req.command {
                DebugCommand::Restart => {
                    let _ = req.reply.send(DebugResponse::Ok);
                    return true;
                }
                DebugCommand::Quit => {
                    let _ = req.reply.send(DebugResponse::Ok);
                    return false;
                }
                _ => {
                    let _ = req.reply.send(DebugResponse::Error(
                        "program exited; use restart or q".to_string(),
                    ));
                }
            }
        }
        false
    }
    /// Prepare an attached debugger for another run in the same VM. The
    /// transport and user breakpoints survive; instruction history belongs to
    /// the previous generation of chunks and must be discarded.
    pub(crate) fn rearm_after_restart(&mut self) {
        self.mode = RunMode::PauseNext;
        self.armed = false;
        self.prelude_skip_done = false;
        self.defer_cmds_until_pause = false;
        self.pending_reply = None;
        self.recent_instructions.clear();
        self.host_names.clear();
        self.sampled_instructions = 0;
        let mut report = self.report.lock().unwrap();
        if report.process_started.is_some() {
            report.process_started = Some(Instant::now());
        }
        report.milestones.clear();
        report.first_output = None;
        report.last_output = None;
        report.paused = Duration::ZERO;
        report.live.instructions.store(0, Ordering::Relaxed);
        report.live.chunk.store(0, Ordering::Relaxed);
        report.live.ip.store(0, Ordering::Relaxed);
        report.live.op.store(0, Ordering::Relaxed);
        report.active.clear();
        report.completed.clear();
        report.active_host = None;
        report.location = None;
        report.instructions = 0;
        report.function_samples.clear();
        report.host_samples.clear();
        report.host_sample_tick = 0;
        report.line_samples.clear();
    }
    pub(crate) fn recent_instruction_report(&self, vm: &VM) -> String {
        self.recent_instructions.iter().map(|(chunk_index, ip, op)| {
            let chunk = vm.chunks.get(*chunk_index);
            let name = chunk.map_or("<unknown>", |chunk| chunk.name.as_str());
            let line = chunk.and_then(|chunk| chunk.get_line(*ip)).unwrap_or(0);
            format!("  {name}@{ip} {op:?} line {line}")
        }).collect::<Vec<_>>().join("\n")
    }

    /// Create a debugger. `pause_on_entry` stops before the first instruction so
    /// the client can set breakpoints (like `node --inspect-brk`).
    pub fn new(
        cmd_rx: Receiver<DebugRequest>,
        evt_tx: Sender<DebugEvent>,
        pause_on_entry: bool,
    ) -> Self {
        let report = Arc::new(Mutex::new(DebugReport::default()));
        let live = report.lock().unwrap().live.clone();
        Debugger {
            report,
            host_names: HashMap::new(),
            live,
            sampled_instructions: 0,
            recent_instructions: VecDeque::with_capacity(32),
            cmd_rx,
            evt_tx,
            breakpoints: Vec::new(),
            next_bp_id: 1,
            watchpoints: Vec::new(),
            next_wp_id: 1,
            break_on_throw: false,
            break_on_uncaught: false,
            break_host: None,
            mode: if pause_on_entry {
                RunMode::PauseNext
            } else {
                RunMode::Running
            },
            stream_opcodes: false,
            trace_global_inits: false,
            break_global_inits: Vec::new(),
            watches: Vec::new(),
            armed: false,
            skip_system: true,
            prelude_skip_done: false,
            pending_reply: None,
            defer_cmds_until_pause: false,
        }
    }

    /// Observe instantiation outside the bytecode dispatch loop. The loader
    /// calls this before each initializer, so the normal REPL can break and
    /// inspect the live VM without an external debugger.
    pub(crate) fn on_global_init(
        &mut self,
        vm: &mut VM,
        chunk_index: usize,
        name: &str,
        init: &crate::chunk::ConstExpr,
    ) -> Result<(), VMError> {
        let breaks = self
            .break_global_inits
            .iter()
            .any(|wanted| wanted == "*" || wanted == name);
        if self.trace_global_inits || breaks {
            let current = vm.global(name).map(render_value).unwrap_or_else(|| "<missing>".into());
            let assigned = vm
                .global_index
                .get(name)
                .and_then(|&i| vm.globals_assigned.get(i as usize))
                .copied()
                .unwrap_or(false);
            let _ = self.evt_tx.send(DebugEvent::Log {
                message: format!(
                    "global-init before name={name} expr={init:?} current={current} assigned={assigned}"
                ),
            });
        }
        if breaks {
            self.enter_pause(vm, chunk_index, 0, PauseReason::Interrupt)?;
        }
        Ok(())
    }

    /// The ordinary entry pause happens at the first instruction, after
    /// instantiation. Pause earlier when the entry module has initializers so
    /// `bgi` and `trace-global-inits` can be configured before they run.
    pub(crate) fn before_global_inits(
        &mut self,
        vm: &mut VM,
        chunk_index: usize,
    ) -> Result<(), VMError> {
        if vm.frames.is_empty() && !self.prelude_skip_done {
            self.prelude_skip_done = true;
            self.enter_pause(vm, chunk_index, 0, PauseReason::Entry)?;
        }
        Ok(())
    }

    pub(crate) fn on_global_init_result(&mut self, name: &str, applied: bool, value: &crate::Value) {
        if self.trace_global_inits
            || self
                .break_global_inits
                .iter()
                .any(|wanted| wanted == "*" || wanted == name)
        {
            let action = if applied { "applied" } else { "skipped (already assigned)" };
            let _ = self.evt_tx.send(DebugEvent::Log {
                message: format!("global-init {action} name={name} value={}", render_value(value)),
            });
        }
    }

    /// Called by the dispatch loop at each instruction boundary (gated by
    /// `VM::instrumented`). `ip` is the offset of the instruction about to run.
    pub fn on_instruction(&mut self, vm: &mut VM, ip: usize, op: Op) -> Result<(), VMError> {
        if vm.frames.is_empty() {
            return Ok(());
        }
        let chunk_index = vm.frames.last().unwrap().chunk_index;
        self.sampled_instructions += 1;
        // A free-running debugger only needs to poll for commands every 64
        // instructions. Keep pauses precise once requested, while avoiding
        // four atomics and recent-history bookkeeping on every ordinary op.
        let free_running = matches!(self.mode, RunMode::Running)
            && self.breakpoints.is_empty()
            && self.watchpoints.is_empty()
            && !self.break_on_throw
            && !self.break_on_uncaught
            && self.break_host.is_none()
            && !self.stream_opcodes;
        if free_running && self.sampled_instructions % 64 != 0 {
            return Ok(());
        }
        if self.recent_instructions.len() == 32 {
            self.recent_instructions.pop_front();
        }
        self.recent_instructions.push_back((chunk_index, ip, op));
        self.live.chunk.store(chunk_index, Ordering::Relaxed);
        self.live.ip.store(ip, Ordering::Relaxed);
        self.live.op.store(op.0, Ordering::Relaxed);
        self.live.instructions.store(self.sampled_instructions, Ordering::Relaxed);
        if self.sampled_instructions == 1 || self.sampled_instructions % 4096 == 0 {
            let name = vm.chunks.get(chunk_index).map_or("<unknown>", |chunk| chunk.name.as_str());
            let mut report = self.report.lock().unwrap();
            report.location = Some(format!("{name}@{ip} (depth {})", vm.frames.len()));
            report.instructions = self.sampled_instructions;
            *report.function_samples.entry(name.to_string()).or_default() += 1;
            if let Some(line) = vm.chunks.get(chunk_index).and_then(|chunk| chunk.get_line(ip)) {
                *report.line_samples.entry(format!("{name}:{line}")).or_default() += 1;
            }
        }

        // 0. Prelude skip (once, on the very first instruction). If the entry
        //    pause is requested and the script chunk marked where user code
        //    begins, convert the entry pause into a one-shot stop at that offset
        //    and run there — so the first pause lands in the user's own code,
        //    not ~200k instructions deep in the runtime prelude. `skip-system
        //    off` disables this. Explicit breakpoints inside the prelude still
        //    fire (they're checked independently).
        if !self.prelude_skip_done {
            self.prelude_skip_done = true;
            if self.skip_system && matches!(self.mode, RunMode::PauseNext) && !self.armed {
                if let Some(off) = vm.chunks.first().and_then(|c| c.user_code_offset) {
                    self.push_breakpoint(0, off as usize, None, None, /* one_shot */ true);
                    self.mode = RunMode::Running;
                    // Until we reach that first user-code pause we are auto-running
                    // ~200k prelude instructions. Do NOT service client commands
                    // during that window — a piped command stream would otherwise
                    // be consumed mid-run (a `c` absorbed by the prelude-skip stop).
                    // Hold them in the channel so the first batch is processed, in
                    // order, at the first real pause (`enter_pause` clears this).
                    self.defer_cmds_until_pause = true;
                }
            }
        }

        // In free running mode with no active stop conditions, only the
        // command channel needs polling. Keep the live frame/IP and recent
        // instruction ring exact above, but amortize the channel and
        // breakpoint machinery over 64 instructions. An asynchronous pause
        // is still serviced within a tiny fraction of a millisecond.
        // 1. Live opcode stream (filterable VYBE_TRACE replacement).
        if self.stream_opcodes {
            let name = vm
                .chunks
                .get(chunk_index)
                .map(|c| c.name.clone())
                .unwrap_or_default();
            let _ = self.evt_tx.send(DebugEvent::Opcode {
                chunk: name,
                ip,
                op: op.wasm_name().to_string(),
                stack_depth: vm.stack.len(),
            });
        }

        // 2. Service any async requests that arrived while running (pause, set-bp,
        //    inspect). Non-blocking; may flip mode so step 3 stops. Suppressed
        //    during the pre-first-pause prelude-skip auto-run (see step 0) so a
        //    piped command batch isn't consumed mid-run.
        if !matches!(self.mode, RunMode::Paused) && !self.defer_cmds_until_pause {
            while let Ok(req) = self.cmd_rx.try_recv() {
                let ctrl = self.handle_command(vm, req.command, chunk_index, ip);
                match ctrl.flow {
                    Flow::Stay => {
                        let _ = req.reply.send(ctrl.response);
                    }
                    Flow::Resume { mode, await_stop } => {
                        self.mode = mode;
                        // Defer the reply until the next stop (see `enter_pause`);
                        // reply now only for non-awaiting resumes (detach/pause).
                        if await_stop {
                            self.pending_reply = Some(req.reply);
                        } else {
                            let _ = req.reply.send(ctrl.response);
                        }
                    }
                    Flow::Quit => {
                        let _ = req.reply.send(ctrl.response);
                        return Err(VMError::new("__debug_quit__"));
                    }
                    Flow::Restart => {
                        let _ = req.reply.send(ctrl.response);
                        return Err(VMError::new("__debug_restart__"));
                    }
                }
            }
        }

        // 3. Decide whether to stop here (with side effects: hit counts,
        //    logpoints, one-shot removal, watchpoint snapshots).
        let host_stop = self.break_host.as_ref().and_then(|(wanted, bad_type_only)| {
            if op != Op::CALL { return None; }
            if *bad_type_only {
                let valid_operand = if wanted == "wasm:js-number:toF64" {
                    matches!(vm.stack.last(), Some(
                        crate::Value::I32(_) | crate::Value::I64(_)
                        | crate::Value::F32(_) | crate::Value::F64(_)
                    ))
                } else {
                    matches!(vm.stack.last(), Some(crate::Value::String(_)))
                };
                if valid_operand { return None; }
            }
            let chunk = vm.chunks.get(chunk_index)?;
            let bytes = chunk.code.get(ip + 4..ip + 6)?;
            let import_idx = u16::from_be_bytes([bytes[0], bytes[1]]) as usize;
            let owner = vm.chunk_import_owner.get(chunk_index).copied().unwrap_or(chunk_index);
            let import = chunk.imports.get(import_idx)
                .or_else(|| vm.chunks.get(owner)?.imports.get(import_idx))?;
            let name = format!("{}:{}", import.module, import.name);
            (wanted == &name).then_some(PauseReason::HostCall { name })
        });
        if let Some(reason) = host_stop.or_else(|| self.decide_stop(vm, chunk_index, ip, op)) {
            self.enter_pause(vm, chunk_index, ip, reason)?;
        }
        self.armed = true;
        Ok(())
    }

    /// The full pause decision: exception breakpoints, location breakpoints (with
    /// conditions / hit counts / logpoints / one-shot), data watchpoints, then
    /// step/entry/interrupt modes. Mutates breakpoint hit counts + watchpoint
    /// snapshots, so takes `&mut self` and `&mut vm` (eval needs it).
    fn decide_stop(
        &mut self,
        vm: &mut VM,
        chunk_index: usize,
        ip: usize,
        op: Op,
    ) -> Option<PauseReason> {
        // (a) Exception breakpoints — about to execute a throw/rethrow/throw_ref.
        if self.armed && is_throw_op(op) {
            let uncaught = vm.exception_handlers.is_empty();
            if self.break_on_throw || (self.break_on_uncaught && uncaught) {
                return Some(PauseReason::Exception { uncaught });
            }
        }

        // (b) Location breakpoints at this (chunk, offset).
        let matched: Vec<u32> = self
            .breakpoints
            .iter()
            .filter(|b| {
                b.enabled
                    && b.offset == ip
                    && (b.chunk_index == chunk_index
                        || b.function_name.as_deref()
                            == vm.chunks.get(chunk_index).map(|chunk| chunk.name.as_str()))
            })
            .map(|b| b.id)
            .collect();
        for id in matched {
            let (ignore, hit, condition, log, one_shot) = {
                let b = self.breakpoints.iter_mut().find(|b| b.id == id).unwrap();
                b.hit_count += 1;
                (
                    b.ignore_count,
                    b.hit_count,
                    b.condition.clone(),
                    b.log_message.clone(),
                    b.one_shot,
                )
            };
            if hit <= ignore {
                continue; // still within the ignore window
            }
            if let Some(cond) = &condition {
                let locals = gather_frame_locals(vm);
                // Fail open (stop) on a broken condition so the user notices.
                let fires = matches!(vm.debug_eval(cond, &locals), Ok(v) if is_truthy(&v))
                    || vm.debug_eval(cond, &locals).is_err();
                if !fires {
                    continue;
                }
            }
            if let Some(msg) = &log {
                let rendered = self.interpolate_log(vm, msg);
                let _ = self.evt_tx.send(DebugEvent::Log { message: rendered });
                continue; // logpoint: never pauses
            }
            if one_shot {
                self.breakpoints.retain(|b| b.id != id);
            }
            return Some(PauseReason::Breakpoint { id });
        }

        // (c) Data watchpoints — value changed since last observation.
        let wp_ids: Vec<u32> = self.watchpoints.iter().map(|w| w.id).collect();
        for wid in wp_ids {
            let target = self
                .watchpoints
                .iter()
                .find(|w| w.id == wid)
                .unwrap()
                .target
                .clone();
            let current = eval_path(vm, &target).unwrap_or_else(|e| format!("<{e}>"));
            let w = self.watchpoints.iter_mut().find(|w| w.id == wid).unwrap();
            if w.last != current {
                let old = std::mem::replace(&mut w.last, current.clone());
                let _ = self.evt_tx.send(DebugEvent::Log {
                    message: format!("watchpoint #{wid}: {target}  {old} → {current}"),
                });
                return Some(PauseReason::Watchpoint { id: wid });
            }
        }

        // (d) Step / entry / interrupt modes.
        match self.mode {
            RunMode::PauseNext if self.armed => Some(PauseReason::Step),
            RunMode::PauseNext => Some(PauseReason::Entry),
            RunMode::StepOver {
                target_depth,
                fiber,
            } => (vm.cur_fiber_id == fiber && vm.frames.len() <= target_depth)
                .then_some(PauseReason::Step),
            RunMode::StepOut {
                target_depth,
                fiber,
            } => (vm.cur_fiber_id == fiber && vm.frames.len() < target_depth)
                .then_some(PauseReason::Step),
            RunMode::Paused => Some(PauseReason::Interrupt),
            RunMode::Running => None,
        }
    }

    /// Block on the command channel, servicing inspection/control requests, until
    /// a request resumes execution.
    fn enter_pause(
        &mut self,
        vm: &mut VM,
        chunk_index: usize,
        ip: usize,
        reason: PauseReason,
    ) -> Result<(), VMError> {
        self.mode = RunMode::Paused;
        let _pause_time = DebugPauseTime(self.report.clone(), Instant::now());
        let location = location_at(vm, chunk_index, ip);
        let frame_summary = if vm.frames.is_empty() {
            "module initialization (no call frame yet)".to_string()
        } else {
            frame_summary(vm)
        };
        // Evaluate watch expressions against this frame (clone the list first so
        // `vm` isn't aliased by a borrow of `self.watches` during eval).
        let watch_exprs = self.watches.clone();
        let mut watches = Vec::with_capacity(watch_exprs.len());
        for w in &watch_exprs {
            let locals = gather_frame_locals(vm);
            let rendered = match vm.debug_eval(w, &locals) {
                Ok(v) => render_value(&v),
                Err(e) => format!("<{e}>"),
            };
            watches.push((w.clone(), rendered));
        }
        let _ = self.evt_tx.send(DebugEvent::Paused {
            reason,
            location,
            frame_summary,
            watches,
        });

        // We have actually stopped — the prelude-skip auto-run (if any) is over,
        // so resume normal command servicing in `on_instruction`.
        self.defer_cmds_until_pause = false;

        // Release the reply held from the resuming command that got us here
        // (`continue`/`step`/…). This is the whole point of deferral: the client's
        // `send_and_print` was blocked until *now*, so it only reads its next piped
        // command while we are genuinely paused. No reply is pending on the entry
        // pause or a fresh interrupt — that's fine.
        if let Some(reply) = self.pending_reply.take() {
            let _ = reply.send(DebugResponse::Ok);
        }

        loop {
            let req = match self.cmd_rx.recv() {
                Ok(r) => r,
                // Client disconnected — detach and let the program finish.
                Err(_) => {
                    self.mode = RunMode::Running;
                    return Ok(());
                }
            };
            let ctrl = self.handle_command(vm, req.command, chunk_index, ip);
            match ctrl.flow {
                Flow::Stay => {
                    let _ = req.reply.send(ctrl.response);
                    continue;
                }
                Flow::Resume { mode, await_stop } => {
                    self.mode = mode;
                    // Hold the reply until the next stop (flushed above) so the
                    // client blocks here; reply now for non-awaiting resumes.
                    if await_stop {
                        self.pending_reply = Some(req.reply);
                    } else {
                        let _ = req.reply.send(ctrl.response);
                    }
                    let _ = self.evt_tx.send(DebugEvent::Resumed);
                    return Ok(());
                }
                Flow::Quit => {
                    let _ = req.reply.send(ctrl.response);
                    return Err(VMError::new("__debug_quit__"));
                }
                Flow::Restart => {
                    let _ = req.reply.send(ctrl.response);
                    return Err(VMError::new("__debug_restart__"));
                }
            }
        }
    }

    /// Execute a single command, producing a reply and a control-flow decision.
    /// Takes `&mut VM` because `SetVar` writes; all other arms only read.
    fn handle_command(
        &mut self,
        vm: &mut VM,
        command: DebugCommand,
        chunk_index: usize,
        ip: usize,
    ) -> Control {
        use DebugCommand::*;
        match command {
            Continue => Control::resume_await(DebugResponse::Ok, RunMode::Running),
            Detach => Control::resume(DebugResponse::Ok, RunMode::Running),
            StepInto | StepInstruction => {
                Control::resume_await(DebugResponse::Ok, RunMode::PauseNext)
            }
            StepOver => Control::resume_await(
                DebugResponse::Ok,
                RunMode::StepOver {
                    target_depth: vm.frames.len(),
                    fiber: vm.cur_fiber_id,
                },
            ),
            StepOut => Control::resume_await(
                DebugResponse::Ok,
                RunMode::StepOut {
                    target_depth: vm.frames.len(),
                    fiber: vm.cur_fiber_id,
                },
            ),
            // Async interrupt: arrange to stop at the next boundary.
            Pause => Control::resume(DebugResponse::Ok, RunMode::PauseNext),
            Quit => Control {
                response: DebugResponse::Ok,
                flow: Flow::Quit,
            },

            BreakLine {
                chunk,
                line,
                condition,
            } => self.set_line_breakpoint(vm, chunk, line, condition),
            BreakOffset {
                chunk,
                offset,
                condition,
            } => self.set_offset_breakpoint(vm, chunk, offset, condition),
            BreakSourceLine { line, condition } => {
                self.set_source_line_breakpoint(vm, line, condition)
            }
            BreakFunction { name, condition } => {
                // A function name is NOT unique: overrides across a class
                // hierarchy compile to separate chunks that share a name (e.g.
                // three `whoami` chunks for A/B/C). A name-based breakpoint
                // covers every matching chunk, including ones loaded later.
                let matches: Vec<usize> = vm
                    .chunks
                    .iter()
                    .enumerate()
                    .filter(|(_, c)| c.name == *name)
                    .map(|(i, _)| i)
                    .collect();
                let id = self.next_bp_id;
                self.next_bp_id += 1;
                self.breakpoints.push(Breakpoint {
                    id,
                    chunk_index: usize::MAX,
                    function_name: Some(name.clone()),
                    offset: 0,
                    enabled: true,
                    condition,
                    ignore_count: 0,
                    hit_count: 0,
                    log_message: None,
                    one_shot: false,
                });
                let line = matches
                    .first()
                    .and_then(|&ci| vm.chunks.get(ci))
                    .and_then(|chunk| chunk.get_line(0));
                let chunk = if matches.is_empty() {
                    format!("{name} (pending)")
                } else if matches.len() > 1 {
                    format!("{name} (×{})", matches.len())
                } else {
                    name
                };
                Control::stay(DebugResponse::BreakpointSet {
                    id,
                    chunk,
                    offset: 0,
                    line,
                })
            }
            BreakHost { name, bad_type_only } => {
                self.break_host = name.clone().map(|name| (name, bad_type_only));
                Control::stay(DebugResponse::Value(match name {
                    Some(name) => format!("host breakpoint set at {name}{}",
                        if bad_type_only { " (invalid operand)" } else { "" }),
                    None => "host breakpoint cleared".to_string(),
                }))
            }
            Logpoint { line, message } => {
                let Some((actual, targets)) = resolve_source_line(vm, line) else {
                    return Control::stay(DebugResponse::Error("no line information".into()));
                };
                for (ci, off) in targets {
                    self.push_breakpoint(ci, off, None, Some(message.clone()), false);
                }
                Control::stay(DebugResponse::Value(format!("logpoint at line {actual}")))
            }
            RunToLine { line } => {
                let Some((_actual, targets)) = resolve_source_line(vm, line) else {
                    return Control::stay(DebugResponse::Error("no line information".into()));
                };
                for (ci, off) in targets {
                    self.push_breakpoint(ci, off, None, None, /* one_shot */ true);
                }
                // …and resume until it's hit (defer the reply until we stop there).
                Control::resume_await(DebugResponse::Ok, RunMode::Running)
            }
            SetIgnoreCount { id, count } => {
                if let Some(bp) = self.breakpoints.iter_mut().find(|b| b.id == id) {
                    bp.ignore_count = count;
                    Control::stay(DebugResponse::Value(format!(
                        "breakpoint #{id} ignores next {count}"
                    )))
                } else {
                    Control::stay(DebugResponse::Error(format!("no breakpoint #{id}")))
                }
            }
            ExceptionBreak {
                on_throw,
                on_uncaught,
            } => {
                self.break_on_throw = on_throw;
                self.break_on_uncaught = on_uncaught;
                let state = match (on_throw, on_uncaught) {
                    (true, _) => "on every throw",
                    (false, true) => "on uncaught only",
                    (false, false) => "off",
                };
                Control::stay(DebugResponse::Value(format!(
                    "exception breakpoints: {state}"
                )))
            }
            Restart => Control {
                response: DebugResponse::Ok,
                flow: Flow::Restart,
            },
            AddWatchpoint { target } => {
                let last = eval_path(vm, &target).unwrap_or_else(|e| format!("<{e}>"));
                let id = self.next_wp_id;
                self.next_wp_id += 1;
                self.watchpoints.push(Watchpoint {
                    id,
                    target: target.clone(),
                    last: last.clone(),
                });
                Control::stay(DebugResponse::Value(format!(
                    "watchpoint #{id} on {target} (= {last})"
                )))
            }
            ListWatchpoints => {
                if self.watchpoints.is_empty() {
                    Control::stay(DebugResponse::Value("(no watchpoints)".into()))
                } else {
                    let list = self
                        .watchpoints
                        .iter()
                        .map(|w| format!("#{} {} = {}", w.id, w.target, w.last))
                        .collect::<Vec<_>>()
                        .join("\n  ");
                    Control::stay(DebugResponse::Value(list))
                }
            }
            ClearWatchpoints => {
                self.watchpoints.clear();
                Control::stay(DebugResponse::Ok)
            }
            Fibers => Control::stay(DebugResponse::Fibers(fiber_list(vm))),
            ListBreakpoints => Control::stay(DebugResponse::Breakpoints(self.list_breakpoints(vm))),
            ClearBreakpoints => {
                self.breakpoints.clear();
                Control::stay(DebugResponse::Ok)
            }
            DeleteBreakpoint { id } => {
                let before = self.breakpoints.len();
                self.breakpoints.retain(|b| b.id != id);
                if self.breakpoints.len() < before {
                    Control::stay(DebugResponse::Ok)
                } else {
                    Control::stay(DebugResponse::Error(format!("no breakpoint #{id}")))
                }
            }
            EnableBreakpoint { id, enabled } => {
                if let Some(bp) = self.breakpoints.iter_mut().find(|b| b.id == id) {
                    bp.enabled = enabled;
                    Control::stay(DebugResponse::Ok)
                } else {
                    Control::stay(DebugResponse::Error(format!("no breakpoint #{id}")))
                }
            }

            Print { path } => {
                // Fast structural read for a plain name/path (no compile); fall
                // back to the compiler-faithful evaluator for compound
                // expressions (`x + y`, calls, etc.).
                match eval_path(vm, &path) {
                    Ok(rendered) => Control::stay(DebugResponse::Value(rendered)),
                    Err(struct_err) => {
                        let locals = gather_frame_locals(vm);
                        Control::stay(match vm.debug_eval(&path, &locals) {
                            Ok(v) => DebugResponse::Value(render_value(&v)),
                            // Prefer the eval error, but if eval is unavailable
                            // for this language the structural error is the real,
                            // actionable one ("no field `x`") — surface it instead
                            // of the generic "eval isn't available" noise.
                            Err(e)
                                if e.contains("eval unavailable")
                                    || e.contains("eval isn't available") =>
                            {
                                DebugResponse::Error(struct_err)
                            }
                            Err(e) => DebugResponse::Error(e),
                        })
                    }
                }
            }
            SetVar { name, literal } => Control::stay(match set_var(vm, &name, &literal) {
                Ok(note) => DebugResponse::Value(note),
                Err(e) => DebugResponse::Error(e),
            }),
            Backtrace => Control::stay(DebugResponse::Backtrace(backtrace(vm))),
            Locals { frame } => Control::stay(match locals(vm, frame) {
                Ok(slots) => DebugResponse::Locals(slots),
                Err(e) => DebugResponse::Error(e),
            }),
            OperandStack => Control::stay(DebugResponse::OperandStack(operand_stack(vm))),
            Globals { prefix } => {
                Control::stay(DebugResponse::Globals(globals(vm, prefix.as_deref())))
            }
            TraceGlobalInits { enabled } => {
                self.trace_global_inits = enabled;
                Control::stay(DebugResponse::Value(format!(
                    "global initializer trace {}",
                    if enabled { "on" } else { "off" }
                )))
            }
            BreakGlobalInit { name } => {
                if name == "off" {
                    self.break_global_inits.clear();
                    Control::stay(DebugResponse::Value("global initializer breakpoints cleared".into()))
                } else {
                    if !self.break_global_inits.contains(&name) {
                        self.break_global_inits.push(name.clone());
                    }
                    Control::stay(DebugResponse::Value(format!("break on global initializer `{name}`")))
                }
            }
            Disasm { window, offset } => Control::stay(DebugResponse::Disasm {
                current_ip: offset.unwrap_or(ip),
                lines: disasm_window(vm, chunk_index, offset.unwrap_or(ip), window),
            }),
            ControlTargets { offset } => {
                let mut matches = Vec::new();
                if let Some(chunk) = vm.chunks.get(chunk_index) {
                    for (source, targets) in crate::dispatch::build_block_table(&chunk.code) {
                        let near_end = targets.end_ip.abs_diff(offset) <= 8;
                        let near_else = targets.else_ip.is_some_and(|at| (at + 4).abs_diff(offset) <= 8);
                        if near_end || near_else {
                            let (instruction, _) = crate::debug::disassemble_instruction(chunk, source);
                            matches.push(format!("{source}: {instruction} -> else {:?}, end {}", targets.else_ip.map(|at| at + 4), targets.end_ip));
                        }
                    }
                }
                matches.sort();
                Control::stay(DebugResponse::Value(if matches.is_empty() {
                    format!("no structured-control target within 8 bytes of {offset}")
                } else { matches.join("\n") }))
            }
            Chunks => Control::stay(DebugResponse::Chunks(chunk_list(vm))),
            Frame { frame } => Control::stay(match frame_shape(vm, frame) {
                Ok(shape) => DebugResponse::Frame(shape),
                Err(e) => DebugResponse::Error(e),
            }),
            AddWatch { expr } => {
                self.watches.push(expr.clone());
                Control::stay(DebugResponse::Value(format!(
                    "watching `{expr}` (#{})",
                    self.watches.len()
                )))
            }
            ListWatches => {
                if self.watches.is_empty() {
                    Control::stay(DebugResponse::Value("(no watches)".into()))
                } else {
                    let list = self
                        .watches
                        .iter()
                        .enumerate()
                        .map(|(i, w)| format!("#{} {w}", i + 1))
                        .collect::<Vec<_>>()
                        .join("\n  ");
                    Control::stay(DebugResponse::Value(list))
                }
            }
            ClearWatches => {
                self.watches.clear();
                Control::stay(DebugResponse::Ok)
            }
            Reload => Control::stay(match vm.debug_reload() {
                Ok(report) => DebugResponse::Value(report),
                Err(e) => DebugResponse::Error(e),
            }),
            StreamOpcodes { enabled } => {
                self.stream_opcodes = enabled;
                Control::stay(DebugResponse::Ok)
            }
            SetSkipSystem { enabled } => {
                self.skip_system = enabled;
                Control::stay(DebugResponse::Value(format!(
                    "skip-system {}",
                    if enabled { "on" } else { "off" }
                )))
            }
            FireEvent { control, event } => Control::stay(match vm.fire_event(&control, &event) {
                Ok(v) => {
                    DebugResponse::Value(format!("fired {control}.{event} → {}", render_value(&v)))
                }
                Err(e) => DebugResponse::Error(e),
            }),
        }
    }

    fn set_line_breakpoint(
        &mut self,
        vm: &VM,
        chunk: ChunkRef,
        line: u32,
        condition: Option<String>,
    ) -> Control {
        // If the `chunk` part doesn't name a real chunk, treat it as a source
        // file reference (`foo.js:7`) and break on the line across all chunks.
        let Some(ci) = resolve_chunk(vm, &chunk) else {
            return self.set_source_line_breakpoint(vm, line, condition);
        };
        match resolve_line_to_offset(&vm.chunks[ci], line) {
            Some(offset) => self.install_breakpoint(vm, ci, offset, condition),
            None => Control::stay(DebugResponse::Error(format!(
                "no instruction on line {line} of chunk {}",
                vm.chunks[ci].name
            ))),
        }
    }

    /// Break on a source line across every chunk (file:line). Slides to the
    /// nearest line that actually has code when the requested line has none
    /// (gdb-style), so any line is breakable. One breakpoint per matching chunk.
    fn set_source_line_breakpoint(
        &mut self,
        vm: &VM,
        line: u32,
        condition: Option<String>,
    ) -> Control {
        let Some((_actual, targets)) = resolve_source_line(vm, line) else {
            return Control::stay(DebugResponse::Error(
                "program has no line information".to_string(),
            ));
        };
        let mut set = Vec::new();
        for (ci, offset) in targets {
            let id = self.push_breakpoint(ci, offset, condition.clone(), None, false);
            set.push(BreakpointInfo {
                id,
                chunk_index: ci,
                chunk_name: vm.chunks[ci].name.clone(),
                offset,
                line: vm.chunks[ci].get_line(offset),
                enabled: true,
            });
        }
        Control::stay(DebugResponse::Breakpoints(set))
    }

    fn set_offset_breakpoint(
        &mut self,
        vm: &VM,
        chunk: ChunkRef,
        offset: usize,
        condition: Option<String>,
    ) -> Control {
        let Some(ci) = resolve_chunk(vm, &chunk) else {
            return Control::stay(DebugResponse::Error(format!("no such chunk: {chunk:?}")));
        };
        self.install_breakpoint(vm, ci, offset, condition)
    }

    fn install_breakpoint(
        &mut self,
        vm: &VM,
        chunk_index: usize,
        offset: usize,
        condition: Option<String>,
    ) -> Control {
        let id = self.push_breakpoint(chunk_index, offset, condition, None, false);
        let line = vm.chunks.get(chunk_index).and_then(|c| c.get_line(offset));
        let chunk = vm
            .chunks
            .get(chunk_index)
            .map(|c| c.name.clone())
            .unwrap_or_default();
        Control::stay(DebugResponse::BreakpointSet {
            id,
            chunk,
            offset,
            line,
        })
    }

    /// Add a breakpoint record and return its id.
    fn push_breakpoint(
        &mut self,
        chunk_index: usize,
        offset: usize,
        condition: Option<String>,
        log_message: Option<String>,
        one_shot: bool,
    ) -> u32 {
        let id = self.next_bp_id;
        self.next_bp_id += 1;
        self.breakpoints.push(Breakpoint {
            id,
            chunk_index,
            function_name: None,
            offset,
            enabled: true,
            condition,
            ignore_count: 0,
            hit_count: 0,
            log_message,
            one_shot,
        });
        id
    }

    /// Interpolate `{expr}` placeholders in a logpoint message using the current
    /// frame (structural read only — cheap, no compile).
    fn interpolate_log(&self, vm: &VM, msg: &str) -> String {
        let mut out = String::new();
        let mut rest = msg;
        while let Some(open) = rest.find('{') {
            out.push_str(&rest[..open]);
            let after = &rest[open + 1..];
            if let Some(close) = after.find('}') {
                let expr = &after[..close];
                match eval_path(vm, expr.trim()) {
                    Ok(v) => out.push_str(&v),
                    Err(_) => {
                        out.push('{');
                        out.push_str(expr);
                        out.push('}');
                    }
                }
                rest = &after[close + 1..];
            } else {
                out.push('{');
                rest = after;
            }
        }
        out.push_str(rest);
        out
    }

    fn list_breakpoints(&self, vm: &VM) -> Vec<BreakpointInfo> {
        self.breakpoints
            .iter()
            .map(|b| BreakpointInfo {
                id: b.id,
                chunk_index: b.chunk_index,
                chunk_name: b.function_name.clone().unwrap_or_else(|| {
                    vm.chunks
                        .get(b.chunk_index)
                        .map(|c| c.name.clone())
                        .unwrap_or_default()
                }),
                offset: b.offset,
                line: vm
                    .chunks
                    .get(b.chunk_index)
                    .and_then(|c| c.get_line(b.offset)),
                enabled: b.enabled,
            })
            .collect()
    }
}

// ─── Control-flow plumbing for command handling ─────────────────────────────

enum Flow {
    Stay,
    /// Resume execution. When `await_stop` is true the command's reply is
    /// *deferred* until the VM next pauses (or the program ends) — this is what
    /// makes `continue`/`step`/`next`/`out`/`run-to` block the client until the
    /// next stop, so a one-shot piped command stream stays in lock-step instead
    /// of racing ahead into a running VM. `await_stop` is false only for commands
    /// that resume and never come back on their own (detach, async pause-request).
    Resume {
        mode: RunMode,
        await_stop: bool,
    },
    Quit,
    Restart,
}

struct Control {
    response: DebugResponse,
    flow: Flow,
}

impl Control {
    fn stay(response: DebugResponse) -> Self {
        Control {
            response,
            flow: Flow::Stay,
        }
    }
    /// Resume and reply immediately (detach / async pause-request — no guaranteed
    /// next stop to defer the reply to).
    fn resume(response: DebugResponse, mode: RunMode) -> Self {
        Control {
            response,
            flow: Flow::Resume {
                mode,
                await_stop: false,
            },
        }
    }
    /// Resume and defer the reply until the next pause — the correct behavior for
    /// stepping/continue so the client blocks until execution actually stops.
    fn resume_await(response: DebugResponse, mode: RunMode) -> Self {
        Control {
            response,
            flow: Flow::Resume {
                mode,
                await_stop: true,
            },
        }
    }
}

// ─── Inspection helpers (read-only against a frozen VM) ─────────────────────

fn resolve_chunk(vm: &VM, chunk: &ChunkRef) -> Option<usize> {
    match chunk {
        ChunkRef::Index(i) => (*i < vm.chunks.len()).then_some(*i),
        ChunkRef::Name(name) => vm.chunks.iter().position(|c| &c.name == name),
    }
}

/// Resolve a source line to `(actual_line, [(chunk, offset)])`, sliding to the
/// nearest line that has instructions when the exact line has none (gdb-style).
/// Prefers the most *specific* chunk whose line range brackets the target (the
/// enclosing function, not the whole `<script>`), so `b 8` lands in the loop
/// rather than an unrelated chunk that happens to share the slid line.
fn resolve_source_line(vm: &VM, target: u32) -> Option<(u32, Vec<(usize, usize)>)> {
    use std::collections::BTreeMap;
    // Per-chunk: line → first instruction offset, in source order.
    let mut per_chunk: Vec<BTreeMap<u32, usize>> = Vec::with_capacity(vm.chunks.len());
    for chunk in vm.chunks.iter() {
        let mut lines: BTreeMap<u32, usize> = BTreeMap::new();
        let mut off = 0;
        while off < chunk.code.len() {
            if let Some(line) = chunk.get_line(off) {
                lines.entry(line).or_insert(off);
            }
            let (_, next) = crate::debug::disassemble_instruction(chunk, off);
            if next <= off {
                break;
            }
            off = next;
        }
        per_chunk.push(lines);
    }

    // Exact line: break in every chunk that has it (a line can legitimately be
    // in more than one function).
    let exact: Vec<(usize, usize)> = per_chunk
        .iter()
        .enumerate()
        .filter_map(|(ci, lines)| lines.get(&target).map(|&off| (ci, off)))
        .collect();
    if !exact.is_empty() {
        return Some((target, exact));
    }

    // Nearest within a single chunk. Score each chunk's best line, preferring
    // chunks whose [min..=max] range contains the target (the enclosing fn),
    // and among those the most specific (smallest range).
    let nearest_in = |lines: &BTreeMap<u32, usize>| -> Option<(u32, usize, u32)> {
        let fwd = lines.range(target..).next();
        let bwd = lines.range(..target).next_back();
        let pick = match (fwd, bwd) {
            (Some((&f, &fo)), Some((&b, &bo))) => {
                if f - target <= target - b {
                    (f, fo)
                } else {
                    (b, bo)
                }
            }
            (Some((&f, &fo)), None) => (f, fo),
            (None, Some((&b, &bo))) => (b, bo),
            (None, None) => return None,
        };
        let dist = pick.0.abs_diff(target);
        Some((pick.0, pick.1, dist))
    };

    let mut best: Option<(bool, u32, u32, usize, usize)> = None; // (contains, span, dist, ci, off)
    for (ci, lines) in per_chunk.iter().enumerate() {
        if lines.is_empty() {
            continue;
        }
        let (&min, _) = lines.iter().next().unwrap();
        let (&max, _) = lines.iter().next_back().unwrap();
        let contains = min <= target && target <= max;
        let span = max - min;
        if let Some((line, off, dist)) = nearest_in(lines) {
            // Rank: containing chunks first, then smaller span, then nearer line.
            let key = (!contains, span, dist);
            let better = best.map_or(true, |(bc, bs, bd, _, _)| key < (!bc, bs, bd));
            if better {
                best = Some((contains, span, dist, ci, off));
                // stash the chosen line via off lookup below
                let _ = line;
            }
        }
    }
    let (_, _, _, ci, off) = best?;
    let actual = vm.chunks[ci].get_line(off).unwrap_or(target);
    Some((actual, vec![(ci, off)]))
}

/// First instruction-start offset whose source line matches `line`.
fn resolve_line_to_offset(chunk: &crate::chunk::Chunk, line: u32) -> Option<usize> {
    let mut offset = 0;
    while offset < chunk.code.len() {
        if chunk.get_line(offset) == Some(line) {
            return Some(offset);
        }
        let (_, next) = crate::debug::disassemble_instruction(chunk, offset);
        if next <= offset {
            break;
        }
        offset = next;
    }
    None
}

fn location_at(vm: &VM, chunk_index: usize, ip: usize) -> Location {
    Location {
        chunk_index,
        chunk_name: vm
            .chunks
            .get(chunk_index)
            .map(|c| c.name.clone())
            .unwrap_or_default(),
        ip,
        line: vm.chunks.get(chunk_index).and_then(|c| c.get_line(ip)),
    }
}

fn frame_summary(vm: &VM) -> String {
    let depth = vm.frames.len();
    let top = vm
        .frames
        .last()
        .map(|f| {
            vm.chunks
                .get(f.chunk_index)
                .map(|c| c.name.as_str())
                .unwrap_or("?")
        })
        .unwrap_or("?");
    format!("{depth} frame(s), in {top}")
}

fn backtrace(vm: &VM) -> Vec<FrameInfo> {
    vm.frames
        .iter()
        .rev()
        .enumerate()
        .map(|(depth, f)| FrameInfo {
            depth,
            chunk_index: f.chunk_index,
            chunk_name: vm
                .chunks
                .get(f.chunk_index)
                .map(|c| c.name.clone())
                .unwrap_or_default(),
            ip: f.ip,
            line: vm
                .chunks
                .get(f.chunk_index)
                .and_then(|c| c.get_line(f.ip.saturating_sub(1))),
        })
        .collect()
}

/// Locals for frame `frame_idx` counting from the top (0 = innermost). Shows the
/// frame's USER variables — slots the compiler gave a source name that isn't an
/// internal temp (`__`-prefixed). Compiler scratch (of which a chunk can have
/// hundreds) is hidden so the view stays the program's actual variables.
/// The frame's declared layout beside its actual contents — see [`FrameShape`].
///
/// Unlike [`locals`] this filters NOTHING: a slot the compiler owns is the slot
/// most likely to be wrong, and naming its owner is what makes a collision
/// legible ("`__capture_0` holds a number") instead of invisible.
fn frame_shape(vm: &VM, frame_idx: usize) -> Result<FrameShape, String> {
    let n = vm.frames.len();
    if frame_idx >= n {
        return Err(format!("no frame #{frame_idx} (have {n})"));
    }
    let real = n - 1 - frame_idx;
    let base = vm.frames[real].base;
    let chunk_index = vm.frames[real].chunk_index;
    let chunk = vm
        .chunks
        .get(chunk_index)
        .ok_or_else(|| format!("frame #{frame_idx} names chunk {chunk_index}, which is absent"))?;

    let hard_end = vm
        .frames
        .get(real + 1)
        .map(|f| f.base)
        .unwrap_or(vm.stack.len());

    // Last binding wins, matching how the compiler reuses a slot across scopes.
    let mut bindings: std::collections::HashMap<u16, &crate::chunk::LocalName> =
        std::collections::HashMap::new();
    for entry in chunk.local_names.iter() {
        bindings.insert(entry.slot, entry);
    }

    let capture_hi = chunk.capture_base as usize + chunk.capture_count as usize;
    let slots = (0..chunk.local_count as usize)
        .map(|index| {
            let binding = bindings.get(&(index as u16)).copied();
            let name = binding.map(|b| b.name.clone());
            // Origin is recorded where the binding is created, so this asks
            // the compiler rather than guessing from the spelling.
            let owner = if chunk.capture_count > 0
                && index >= chunk.capture_base as usize
                && index < capture_hi
            {
                SlotOwner::Capture
            } else if binding.is_some_and(|b| b.is_source()) {
                SlotOwner::User
            } else {
                SlotOwner::Compiler
            };
            let abs = base + index;
            let live = abs < hard_end && abs < vm.stack.len();
            FrameSlot {
                index,
                name,
                owner,
                value: if live {
                    render_value(&vm.stack[abs])
                } else {
                    "<unset>".into()
                },
                live,
            }
        })
        .collect();

    Ok(FrameShape {
        frame: frame_idx,
        chunk_index,
        chunk_name: chunk.name.clone(),
        arity: chunk.arity,
        local_count: chunk.local_count,
        scratch_high_water: chunk.scratch_high_water,
        capture_base: chunk.capture_base,
        capture_count: chunk.capture_count,
        base,
        slots,
    })
}

fn locals(vm: &VM, frame_idx: usize) -> Result<Vec<SlotInfo>, String> {
    let n = vm.frames.len();
    if frame_idx >= n {
        return Err(format!("no frame #{frame_idx} (have {n})"));
    }
    let real = n - 1 - frame_idx;
    let base = vm.frames[real].base;
    let Some(chunk) = vm.chunks.get(vm.frames[real].chunk_index) else {
        return Ok(Vec::new());
    };
    let hard_end = vm
        .frames
        .get(real + 1)
        .map(|f| f.base)
        .unwrap_or(vm.stack.len());
    // One entry per slot (last binding wins), user names only, in slot order.
    let mut seen = std::collections::HashSet::new();
    let mut slots: Vec<SlotInfo> = Vec::new();
    for entry in chunk.local_names.iter().rev() {
        if !entry.is_source() {
            continue; // compiler-owned; `frame` shows these, `locals` does not
        }
        let (slot, name) = (&entry.slot, &entry.name);
        if !seen.insert(*slot) {
            continue;
        }
        let idx = base + *slot as usize;
        if idx >= hard_end || idx >= vm.stack.len() {
            continue;
        }
        slots.push(SlotInfo {
            index: *slot as usize,
            name: Some(name.clone()),
            value: render_value(&vm.stack[idx]),
        });
    }
    slots.sort_by_key(|s| s.index);
    Ok(slots)
}

/// The operand stack of the innermost frame: everything the current chunk has
/// pushed above its reserved locals/scratch region. (The full VM stack also
/// holds every outer frame's saved operands, which would bury the useful view.)
fn operand_stack(vm: &VM) -> Vec<String> {
    let Some(frame) = vm.frames.last() else {
        return vm.stack.iter().map(render_value).collect();
    };
    let local_count = vm
        .chunks
        .get(frame.chunk_index)
        .map(|c| c.local_count as usize)
        .unwrap_or(0);
    let start = (frame.base + local_count).min(vm.stack.len());
    vm.stack[start..].iter().map(render_value).collect()
}

fn globals(vm: &VM, prefix: Option<&str>) -> Vec<(String, String)> {
    let mut out: Vec<(String, String)> = vm
        .globals_by_name()
        .into_iter()
        .filter(|(k, _)| prefix.map(|p| k.starts_with(p)).unwrap_or(true))
        .map(|(k, v)| {
            let assigned = vm
                .global_index
                .get(&k)
                .and_then(|&i| vm.globals_assigned.get(i as usize))
                .copied()
                .unwrap_or(false);
            let value = render_value(&v);
            (k, if assigned { value } else { format!("{value} [unassigned]") })
        })
        .collect();
    out.sort_by(|a, b| a.0.cmp(&b.0));
    out
}

fn disasm_window(vm: &VM, chunk_index: usize, current_ip: usize, window: usize) -> Vec<DisasmLine> {
    let Some(chunk) = vm.chunks.get(chunk_index) else {
        return Vec::new();
    };
    // Collect all instruction starts, then window around the current ip.
    let mut instrs: Vec<usize> = Vec::new();
    let mut offset = 0;
    while offset < chunk.code.len() {
        instrs.push(offset);
        let (_, next) = crate::debug::disassemble_instruction(chunk, offset);
        if next <= offset {
            break;
        }
        offset = next;
    }
    let cur_idx = instrs.partition_point(|&o| o <= current_ip).saturating_sub(1);
    let start = cur_idx.saturating_sub(window);
    let end = (cur_idx + window + 1).min(instrs.len());
    instrs[start..end]
        .iter()
        .map(|&o| {
            let (text, _) = crate::debug::disassemble_instruction(chunk, o);
            DisasmLine {
                offset: o,
                text,
                is_current: o == current_ip,
                line: chunk.get_line(o),
            }
        })
        .collect()
}

/// One summary line per live fiber: the current fiber plus every suspended one
/// (continuations + event-loop waiters + queued resume tasks).
fn fiber_list(vm: &VM) -> Vec<String> {
    let name_of = |ci: usize| {
        vm.chunks
            .get(ci)
            .map(|c| c.name.as_str())
            .unwrap_or("?")
            .to_string()
    };
    let mut out = Vec::new();
    let cur_top = vm
        .frames
        .last()
        .map(|f| name_of(f.chunk_index))
        .unwrap_or_else(|| "?".into());
    out.push(format!(
        "* fiber #{} (current) — {} frame(s), in {}",
        vm.cur_fiber_id,
        vm.frames.len(),
        cur_top
    ));
    for (n, ac) in vm.debug_suspended_fibers().into_iter().enumerate() {
        let top =
            ac.0.last()
                .map(|&ci| name_of(ci))
                .unwrap_or_else(|| "?".into());
        out.push(format!(
            "  suspended #{n} ({}) — {} frame(s), in {}",
            ac.1,
            ac.0.len(),
            top
        ));
    }
    out
}

fn chunk_list(vm: &VM) -> Vec<ChunkInfo> {
    vm.chunks
        .iter()
        .enumerate()
        .map(|(index, c)| ChunkInfo {
            index,
            name: c.name.clone(),
            arity: c.arity,
            code_len: c.code.len(),
        })
        .collect()
}

/// Bounded rendering of a value so a huge graph can't flood the wire.
fn render_value(v: &crate::Value) -> String {
    // Snapshot first: a map can contain itself, so never format a child
    // while holding the parent's lock. Keep nested objects shallow.
    let entries = if let crate::Value::Object(object) = v {
        let object = object.lock().unwrap();
        if let crate::value::ObjectKind::Map(entries) = &object.kind {
            Some((entries.len(), entries.iter().take(6)
                .map(|(key, value)| (key.clone(), value.clone())).collect::<Vec<_>>()))
        } else { None }
    } else { None };
    let s = if let Some((len, entries)) = entries {
        let item = |value: &crate::Value| match value {
            crate::Value::Object(_) => "[object]".to_string(),
            crate::Value::String(value) => format!("{value:?}"),
            _ => value.to_string(),
        };
        format!("Map({len}) {{{}{}}}", entries.iter()
            .map(|(key, value)| format!("{}: {}", item(key), item(value)))
            .collect::<Vec<_>>().join(", "), if len > entries.len() { ", …" } else { "" })
    } else { format!("{}", v) };
    const CAP: usize = 200;
    if s.len() > CAP {
        let mut end = CAP;
        while !s.is_char_boundary(end) { end -= 1; }
        format!("{}…", &s[..end])
    } else {
        s
    }
}

/// True for the throw-family core opcodes (`throw` 0x08, `rethrow` 0x09,
/// `throw_ref` 0x0A) — used by exception breakpoints.
fn is_throw_op(op: Op) -> bool {
    op.group() == 0 && matches!(op.sub(), 0x08 | 0x09 | 0x0A)
}

/// Truthiness for conditional breakpoints. The condition's operator semantics
/// (e.g. `a > 5`) were already applied faithfully by the real compiler; this
/// only maps the resulting value to stop/skip. Comparisons yield a bool, the
/// common case.
fn is_truthy(v: &crate::Value) -> bool {
    use crate::Value::*;
    match v {
        Bool(b) => *b,
        Null | TypedNull(_) | Undefined => false,
        I32(n) => *n != 0,
        I64(n) => *n != 0,
        F32(n) => *n != 0.0,
        F64(n) => *n != 0.0,
        String(s) => !s.is_empty(),
        _ => true,
    }
}

// ─── Variable read/write by name (Pass 2 — no expression semantics) ─────────

/// One `.field` / `[index]` step in a `Print` path.
enum Access {
    Field(String),
    Index(usize),
}

/// Collect the innermost frame's locals as `(name, value)` for injection into
/// the eval mini-VM. One entry per slot — the last (currently-live) binding
/// wins when sibling blocks reused a slot.
fn gather_frame_locals(vm: &VM) -> Vec<(String, crate::Value)> {
    let Some(frame) = vm.frames.last() else {
        return Vec::new();
    };
    let Some(chunk) = vm.chunks.get(frame.chunk_index) else {
        return Vec::new();
    };
    let mut seen_slots = std::collections::HashSet::new();
    let mut out = Vec::new();
    for entry in chunk.local_names.iter().rev() {
        if !entry.is_source() {
            continue; // compiler-owned — never part of user eval scope
        }
        let (slot, name) = (&entry.slot, &entry.name);
        if !seen_slots.insert(*slot) {
            continue;
        }
        let idx = frame.base + *slot as usize;
        if let Some(v) = vm.stack.get(idx) {
            out.push((name.clone(), v.clone()));
        }
    }
    out
}

/// Read a variable's live value in the innermost frame: a local (resolved via
/// the compiler's debug names) or, failing that, a global.
fn read_named_value(vm: &VM, name: &str) -> Option<crate::Value> {
    let alternate = if let Some(bare) = name.strip_prefix('$') {
        bare.to_string()
    } else {
        format!("${name}")
    };
    for candidate in [name, alternate.as_str()] {
        if let Some(frame) = vm.frames.last() {
            if let Some(chunk) = vm.chunks.get(frame.chunk_index) {
                if let Some(entry) = chunk.local_names.iter().rev().find(|e| e.name == candidate) {
                    let idx = frame.base + entry.slot as usize;
                    if let Some(v) = vm.stack.get(idx) {
                        return Some(v.clone());
                    }
                }
            }
        }
        if let Some(value) = vm.global(candidate) {
            return Some(value.clone());
        }
    }
    // PHP's top-level variables live in compiler-managed global cells rather
    // than the source spelling. Keep this mapping at the debugger boundary so
    // a structural `p $name` never falls into expression evaluation.
    let bare = name.strip_prefix('$').unwrap_or(name);
    vm.global(&format!("__php_var_{bare}")).cloned()
}

/// Parse `base.field[0].other` (or PHP `$base->field`) into a base name and
/// accessors. This is structural drill-down, NOT expression evaluation.
fn parse_path(path: &str) -> Result<(String, Vec<Access>), String> {
    let bytes = path.as_bytes();
    let mut i = 0;
    if bytes.first() == Some(&b'$') {
        i = 1;
    }
    let start = i;
    while i < bytes.len() && (bytes[i] == b'_' || bytes[i].is_ascii_alphanumeric()) {
        i += 1;
    }
    if i == start {
        return Err(format!("`{path}` is not a variable name"));
    }
    // Keep a leading `$`: PHP locals and globals can coexist with a bare
    // callable of the same name, and the debugger must prefer the variable.
    let base = path[..i].to_string();
    let mut accessors = Vec::new();
    while i < bytes.len() {
        match bytes[i] {
            b'.' | b'-' => {
                let php_arrow = bytes[i] == b'-';
                i += 1;
                if php_arrow {
                    if bytes.get(i) != Some(&b'>') {
                        return Err("expected `->` after `-`".into());
                    }
                    i += 1;
                }
                let s = i;
                while i < bytes.len() && (bytes[i] == b'_' || bytes[i].is_ascii_alphanumeric()) {
                    i += 1;
                }
                if i == s {
                    return Err("expected a field name after property access".into());
                }
                accessors.push(Access::Field(path[s..i].to_string()));
            }
            b'[' => {
                i += 1;
                let s = i;
                while i < bytes.len() && bytes[i] != b']' {
                    i += 1;
                }
                if i >= bytes.len() {
                    return Err("unterminated `[`".into());
                }
                let idx: usize = path[s..i]
                    .trim()
                    .parse()
                    .map_err(|_| "index must be an integer")?;
                accessors.push(Access::Index(idx));
                i += 1; // consume ']'
            }
            other => return Err(format!("unexpected `{}` in path", other as char)),
        }
    }
    Ok((base, accessors))
}

/// One navigation step into an object property or array element. Pure read; no
/// prototype chain, no coercion — a debugger drill-down.
fn navigate(vm: &VM, val: &crate::Value, acc: &Access) -> Result<crate::Value, String> {
    let crate::Value::Object(o) = val else {
        return Err(format!("`{}` is not indexable", render_value(val)));
    };
    if let Access::Field(field) = acc {
        let obj = o.lock().unwrap();
        if let crate::value::ObjectKind::Map(entries) = &obj.kind {
            let key = crate::Value::String(std::sync::Arc::from(field.as_str()));
            if let Some(value) = entries.get(&key) {
                return Ok(value.clone());
            }
        }
        if let Some(value) = obj.properties.get(field).cloned() {
            return Ok(value);
        }
        drop(obj);
        return vm.resolve_property(val, field).map_err(|e| e.to_string());
    }
    let obj = o.lock().unwrap();
    match acc {
        Access::Field(_) => unreachable!(),
        Access::Index(i) => match &obj.kind {
            crate::value::ObjectKind::Map(entries) => entries
                .get(&crate::Value::I64(*i as i64))
                .or_else(|| entries.get(&crate::Value::String(std::sync::Arc::from(i.to_string()))))
                .cloned().ok_or_else(|| format!("key {i} not found")),
            crate::value::ObjectKind::Array(elems) => elems
                .get(*i)
                .cloned()
                .ok_or_else(|| format!("index {i} out of range")),
            _ => obj
                .fields
                .get(*i)
                .cloned()
                .ok_or_else(|| format!("index {i} out of range")),
        },
    }
}

/// `Print`: resolve a path to a value and render it.
fn eval_path(vm: &VM, path: &str) -> Result<String, String> {
    let (base, accessors) = parse_path(path.trim())?;
    let mut val =
        read_named_value(vm, &base).ok_or_else(|| format!("no variable `{base}` in scope"))?;
    for acc in &accessors {
        val = navigate(vm, &val, acc)?;
    }
    Ok(render_value(&val))
}

/// Parse a literal for `SetVar`: null / undefined / bool / int / float / string
/// (quoted or bare). No expressions.
fn parse_literal(s: &str) -> crate::Value {
    use std::sync::Arc;
    let t = s.trim();
    match t {
        "null" => return crate::Value::Null,
        "undefined" => return crate::Value::Undefined,
        "true" => return crate::Value::Bool(true),
        "false" => return crate::Value::Bool(false),
        _ => {}
    }
    if let Ok(i) = t.parse::<i32>() {
        return crate::Value::I32(i);
    }
    if let Ok(n) = t.parse::<f64>() {
        return crate::Value::F64(n);
    }
    let unquoted = if (t.starts_with('"') && t.ends_with('"') && t.len() >= 2)
        || (t.starts_with('\'') && t.ends_with('\'') && t.len() >= 2)
    {
        &t[1..t.len() - 1]
    } else {
        t
    };
    crate::Value::String(Arc::from(unquoted))
}

/// `SetVar`: write a literal into a named local (innermost frame) or global.
fn set_var(vm: &mut VM, name: &str, literal: &str) -> Result<String, String> {
    let new_val = parse_literal(literal);
    // Use the same structural paths as `p`, so request inputs and nested
    // runtime state can be adjusted without compiling a diagnostic script.
    if let Ok((base, accessors)) = parse_path(name) {
        if let Some((last, parents)) = accessors.split_last() {
            let mut parent = read_named_value(vm, &base)
                .ok_or_else(|| format!("no variable `{base}` in scope"))?;
            for access in parents {
                parent = navigate(vm, &parent, access)?;
            }
            let old = navigate(vm, &parent, last).unwrap_or(crate::Value::Undefined);
            let crate::Value::Object(object) = parent else {
                return Err("assignment parent is not an object or array".into());
            };
            let mut object = object.lock().unwrap();
            match (last, &mut object.kind) {
                (Access::Field(field), crate::value::ObjectKind::Map(entries)) => {
                    entries.insert(crate::Value::String(field.as_str().into()), new_val.clone());
                }
                (Access::Index(index), crate::value::ObjectKind::Map(entries)) => {
                    entries.insert(crate::Value::I64(*index as i64), new_val.clone());
                }
                (Access::Index(index), crate::value::ObjectKind::Array(elements)) => {
                    let target = elements.get_mut(*index).ok_or("index out of range")?;
                    *target = new_val.clone();
                }
                (Access::Field(field), _) => object.set(field.clone(), new_val.clone()),
                (Access::Index(index), _) => {
                    let target = object.fields.get_mut(*index).ok_or("index out of range")?;
                    *target = new_val.clone();
                }
            }
            return Ok(format!("{name}: {} → {}", render_value(&old), render_value(&new_val)));
        }
    }
    // Local in the innermost frame (resolved via debug names).
    if let Some(frame) = vm.frames.last() {
        let slot = vm.chunks.get(frame.chunk_index).and_then(|c| {
            c.local_names
                .iter()
                .rev()
                .find(|e| e.name == name)
                .map(|e| e.slot)
        });
        if let Some(slot) = slot {
            let idx = frame.base + slot as usize;
            if idx < vm.stack.len() {
                let old = render_value(&vm.stack[idx]);
                vm.stack[idx] = new_val.clone();
                return Ok(format!("{name}: {old} → {}", render_value(&new_val)));
            }
        }
    }
    if vm.has_global(name) {
        let old = vm.global(name).map(render_value).unwrap_or_default();
        vm.set_global(name, new_val.clone());
        return Ok(format!("{name}: {old} → {}", render_value(&new_val)));
    }
    Err(format!("no variable `{name}` in scope"))
}
