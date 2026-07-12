crate::js_cases! {
    atomics_sub_returns_old_value_and_updates_slot => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
view[0] = 9;
console.log(Atomics.sub(view, 0, 4));
console.log(Atomics.load(view, 0));
"#,
        ["9", "5"]
    };

    atomics_and_applies_bitmask => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
view[0] = 14;
console.log(Atomics.and(view, 0, 10));
console.log(Atomics.load(view, 0));
"#,
        ["14", "10"]
    };

    atomics_or_merges_bits => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
view[0] = 8;
console.log(Atomics.or(view, 0, 3));
console.log(Atomics.load(view, 0));
"#,
        ["8", "11"]
    };

    atomics_xor_toggles_bits => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
view[0] = 15;
console.log(Atomics.xor(view, 0, 10));
console.log(Atomics.load(view, 0));
"#,
        ["15", "5"]
    };

    atomics_wait_returns_not_equal_when_value_differs => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
view[0] = 1;
console.log(Atomics.wait(view, 0, 0, 1));
"#,
        ["not-equal"]
    };

    atomics_wait_times_out_without_notifier => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
view[0] = 1;
console.log(Atomics.wait(view, 0, 1, 0));
"#,
        ["timed-out"]
    };

    atomics_notify_returns_zero_without_waiters => {
        r#"
const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
console.log(Atomics.notify(view, 0));
"#,
        ["0"]
    };
}
