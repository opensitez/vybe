//! SDL input surface — `sdlplan.md` Tier 1.
//!
//! The window layer pushes raw key/mouse events into `GuiState.input_events`
//! (already in SDL's own vocabulary — see `SdlInputEvent`); these host fns
//! drain and sample that state. `SDL_PollEvent`'s struct-filling happens HERE,
//! not in emitted bytecode: the host can mutate the pointee object directly,
//! so the libc emitter stays a plain call instead of a field-copy sequence.
//!
//! `sdlPushEvent` is the real `SDL_PushEvent` API, which also makes the whole
//! queue testable headless: a C test can inject a keypress and poll it back
//! without a window existing.

use std::sync::{Arc, Mutex};

use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

use crate::gui_state::{GuiState, SdlInputEvent};

/// Unwrap a C pointer argument to the object it addresses.
///
/// `&e` on a struct arrives either as the struct object itself, or boxed in a
/// scalar cell `{__ref_kind:"cell", __value}`. Either way the caller wants the
/// inner object.
fn pointee(v: Option<&Value>) -> Option<Arc<Mutex<Object>>> {
    let Some(Value::Object(obj)) = v else {
        return None;
    };
    let inner = {
        let o = obj.lock().unwrap();
        match o.properties.get("__value") {
            Some(Value::Object(inner))
                if matches!(
                    o.properties.get("__ref_kind"),
                    Some(Value::String(kind)) if kind.as_ref() == "cell"
                ) =>
            {
                Some(inner.clone())
            }
            _ => None }
    };
    Some(inner.unwrap_or_else(|| obj.clone()))
}

fn obj_value(o: Object) -> Value {
    Value::Object(Arc::new(Mutex::new(o)))
}

/// Build the nested `SDL_Event` view of one queued event and write it into
/// `target`'s fields: `type`, `key.keysym.{sym,scancode,mod}`,
/// `motion.{x,y}`, `button.{button,x,y}`, `wheel.y`.
fn fill_event(target: &Arc<Mutex<Object>>, evt: &SdlInputEvent) {
    let mut keysym = Object::new();
    keysym
        .properties
        .insert("sym".into(), Value::I32(evt.sym));
    keysym
        .properties
        .insert("scancode".into(), Value::I32(evt.scancode));
    keysym
        .properties
        .insert("mod".into(), Value::I32(evt.mod_state as i32));
    let mut key = Object::new();
    key.properties.insert("keysym".into(), obj_value(keysym));
    key.properties
        .insert("state".into(), Value::I32(if evt.event_type == 0x300 { 1 } else { 0 }));

    let mut motion = Object::new();
    motion.properties.insert("x".into(), Value::I32(evt.x));
    motion.properties.insert("y".into(), Value::I32(evt.y));

    let mut button = Object::new();
    button
        .properties
        .insert("button".into(), Value::I32(evt.button as i32));
    button.properties.insert("x".into(), Value::I32(evt.x));
    button.properties.insert("y".into(), Value::I32(evt.y));

    let mut wheel = Object::new();
    wheel
        .properties
        .insert("y".into(), Value::I32(evt.wheel_y));

    let mut t = target.lock().unwrap();
    t.properties
        .insert("type".into(), Value::I32(evt.event_type as i32));
    t.properties.insert("key".into(), obj_value(key));
    t.properties.insert("motion".into(), obj_value(motion));
    t.properties.insert("button".into(), obj_value(button));
    t.properties.insert("wheel".into(), obj_value(wheel));
}

/// Write an out-parameter: through a cell's `__value`, or as a plain store
/// into index 0 for an array-backed pointer.
fn write_out_param(v: Option<&Value>, n: i32) {
    let Some(Value::Object(obj)) = v else {
        return;
    };
    let mut o = obj.lock().unwrap();
    if matches!(
        o.properties.get("__ref_kind"),
        Some(Value::String(kind)) if kind.as_ref() == "cell"
    ) {
        o.properties.insert("__value".into(), Value::I32(n));
        return;
    }
    if let ObjectKind::Array(elems) = &mut o.kind {
        if let Some(slot) = elems.get_mut(0) {
            *slot = Value::I32(n);
        }
    }
}

pub fn register(vm: &mut VM, gui: Arc<Mutex<GuiState>>) {
    // SDL_PollEvent(SDL_Event *e) → 1 if an event was dequeued, else 0.
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "sdlPollEvent",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let evt = gui.lock().unwrap().input_events.pop_front();
                let Some(evt) = evt else {
                    return Value::I32(0);
                };
                if let Some(target) = pointee(args.first()) {
                    fill_event(&target, &evt);
                }
                Value::I32(1)
            }),
        );
    }

    // SDL_PushEvent(SDL_Event *e) → 1 on success. Real SDL API; also the
    // headless test path for the whole queue.
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "sdlPushEvent",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(target) = pointee(args.first()) else {
                    return Value::I32(0);
                };
                let evt = {
                    let t = target.lock().unwrap();
                    let get_i32 = |o: &Object, k: &str| -> i32 {
                        o.properties.get(k).map(|v| v.as_i32()).unwrap_or(0)
                    };
                    let nested = |o: &Object, k: &str| -> Option<Arc<Mutex<Object>>> {
                        match o.properties.get(k) {
                            Some(Value::Object(inner)) => Some(inner.clone()),
                            _ => None }
                    };
                    let mut evt = SdlInputEvent::empty(get_i32(&t, "type") as u32);
                    if let Some(key) = nested(&t, "key") {
                        let key = key.lock().unwrap();
                        if let Some(keysym) = nested(&key, "keysym") {
                            let keysym = keysym.lock().unwrap();
                            evt.sym = get_i32(&keysym, "sym");
                            evt.scancode = get_i32(&keysym, "scancode");
                            evt.mod_state = get_i32(&keysym, "mod") as u32;
                        }
                    }
                    if let Some(motion) = nested(&t, "motion") {
                        let motion = motion.lock().unwrap();
                        evt.x = get_i32(&motion, "x");
                        evt.y = get_i32(&motion, "y");
                    }
                    if let Some(button) = nested(&t, "button") {
                        let button = button.lock().unwrap();
                        evt.button = get_i32(&button, "button") as u32;
                        evt.x = get_i32(&button, "x");
                        evt.y = get_i32(&button, "y");
                    }
                    if let Some(wheel) = nested(&t, "wheel") {
                        let wheel = wheel.lock().unwrap();
                        evt.wheel_y = get_i32(&wheel, "y");
                    }
                    evt
                };
                gui.lock().unwrap().push_input_event(evt);
                Value::I32(1)
            }),
        );
    }

    // SDL_GetMouseState(int *x, int *y) → held-button mask.
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "sdlGetMouseState",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (x, y, buttons) = {
                    let g = gui.lock().unwrap();
                    (g.mouse_x, g.mouse_y, g.mouse_buttons)
                };
                write_out_param(args.first(), x);
                write_out_param(args.get(1), y);
                Value::I32(buttons as i32)
            }),
        );
    }

    // SDL_GetModState() → KMOD_* mask.
    {
        let gui = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "sdlGetModState",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                Value::I32(gui.lock().unwrap().mod_state as i32)
            }),
        );
    }
}
