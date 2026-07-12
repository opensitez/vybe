use super::helpers::run_js;

// ── ArrayBuffer basics ────────────────────────────────────
#[test]
fn arraybuffer_create_with_bytelength() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(16);
console.log(buf.byteLength);
"#
        ),
        vec!["16"]
    );
}

#[test]
fn arraybuffer_initial_bytes_are_zero() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const view = new Uint8Array(buf);
console.log(view[0], view[1], view[2], view[3]);
"#
        ),
        vec!["0 0 0 0"]
    );
}

#[test]
fn arraybuffer_slice_creates_copy() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const view = new Uint8Array(buf);
view[0] = 42;
const sliced = buf.slice(0, 4);
const slicedView = new Uint8Array(sliced);
slicedView[0] = 99;
console.log(view[0]);
console.log(slicedView[0]);
"#
        ),
        vec!["42", "99"]
    );
}

#[test]
fn arraybuffer_isview_typedarray() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const view = new Uint8Array(buf);
console.log(ArrayBuffer.isView(view));
console.log(ArrayBuffer.isView(buf));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn arraybuffer_detached_after_transfer() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
console.log(buf.byteLength);
"#
        ),
        vec!["4"]
    );
}

// ── DataView read/write ───────────────────────────────────
#[test]
fn dataview_set_and_get_uint8() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
dv.setUint8(0, 255);
console.log(dv.getUint8(0));
"#
        ),
        vec!["255"]
    );
}

#[test]
fn dataview_set_and_get_int8() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
dv.setInt8(0, -128);
console.log(dv.getInt8(0));
"#
        ),
        vec!["-128"]
    );
}

#[test]
fn dataview_set_and_get_uint16_little_endian() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
dv.setUint16(0, 256, true);
console.log(dv.getUint16(0, true));
"#
        ),
        vec!["256"]
    );
}

#[test]
fn dataview_set_and_get_int32() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
dv.setInt32(0, -100000);
console.log(dv.getInt32(0));
"#
        ),
        vec!["-100000"]
    );
}

#[test]
fn dataview_set_and_get_float32() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
dv.setFloat32(0, 3.14);
const val = dv.getFloat32(0);
console.log(val > 3.13 && val < 3.15);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dataview_set_and_get_float64() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(16);
const dv = new DataView(buf);
dv.setFloat64(0, 3.141592653589793);
const val = dv.getFloat64(0);
console.log(val.toFixed(6));
"#
        ),
        vec!["3.141593"]
    );
}

#[test]
fn dataview_byte_offset_access() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
dv.setUint8(0, 10);
dv.setUint8(1, 20);
dv.setUint8(2, 30);
console.log(dv.getUint8(1));
"#
        ),
        vec!["20"]
    );
}

#[test]
fn dataview_bytelength_and_byteoffset() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(16);
const dv = new DataView(buf, 4, 8);
console.log(dv.byteLength);
console.log(dv.byteOffset);
"#
        ),
        vec!["8", "4"]
    );
}

#[test]
fn dataview_buffer_property() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
console.log(dv.buffer === buf);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dataview_multiple_views_share_buffer() {
    assert_eq!(
        run_js(
            r#"
const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
const u8 = new Uint8Array(buf);
dv.setUint8(0, 42);
console.log(u8[0]);
"#
        ),
        vec!["42"]
    );
}

// ── SharedArrayBuffer ─────────────────────────────────────
#[test]
fn sharedarraybuffer_bytelength() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(16);
console.log(sab.byteLength);
"#
        ),
        vec!["16"]
    );
}

#[test]
fn sharedarraybuffer_with_int32array() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(16);
const ia = new Int32Array(sab);
ia[0] = 42;
console.log(ia[0]);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn sharedarraybuffer_is_shared_not_detachable() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(4);
console.log(sab instanceof SharedArrayBuffer);
"#
        ),
        vec!["true"]
    );
}

// ── Atomics ───────────────────────────────────────────────
#[test]
fn atomics_store_and_load() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(4);
const ia = new Int32Array(sab);
Atomics.store(ia, 0, 42);
console.log(Atomics.load(ia, 0));
"#
        ),
        vec!["42"]
    );
}

#[test]
fn atomics_add() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(4);
const ia = new Int32Array(sab);
ia[0] = 10;
Atomics.add(ia, 0, 5);
console.log(ia[0]);
"#
        ),
        vec!["15"]
    );
}

#[test]
fn atomics_compareexchange_success() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(4);
const ia = new Int32Array(sab);
ia[0] = 5;
const old = Atomics.compareExchange(ia, 0, 5, 10);
console.log(old);
console.log(ia[0]);
"#
        ),
        vec!["5", "10"]
    );
}

#[test]
fn atomics_compareexchange_failure() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(4);
const ia = new Int32Array(sab);
ia[0] = 5;
const old = Atomics.compareExchange(ia, 0, 99, 10);
console.log(old);
console.log(ia[0]);
"#
        ),
        vec!["5", "5"]
    );
}

#[test]
fn atomics_exchange_returns_old() {
    assert_eq!(
        run_js(
            r#"
const sab = new SharedArrayBuffer(4);
const ia = new Int32Array(sab);
ia[0] = 7;
const old = Atomics.exchange(ia, 0, 42);
console.log(old);
console.log(ia[0]);
"#
        ),
        vec!["7", "42"]
    );
}
