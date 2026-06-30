//! DataView multi-byte reads/writes — endianness, bounds, and typed access.

crate::js_cases! {
    dataview_get_int8_reads_signed_byte => {
        r#"const b=new ArrayBuffer(2); const v=new DataView(b); v.setUint8(0,255); console.log(v.getInt8(0));"#,
        ["-1"]
    };

    dataview_get_uint8_reads_unsigned_byte => {
        r#"const b=new ArrayBuffer(1); const v=new DataView(b); v.setUint8(0,200); console.log(v.getUint8(0));"#,
        ["200"]
    };

    dataview_get_int16_little_endian => {
        r#"const b=new ArrayBuffer(2); const v=new DataView(b); v.setInt16(0, 0x0102, true); console.log(v.getInt16(0, true));"#,
        ["258"]
    };

    dataview_get_int16_big_endian => {
        r#"const b=new ArrayBuffer(2); const v=new DataView(b); v.setInt16(0, 0x0102, false); console.log(v.getInt16(0, false));"#,
        ["258"]
    };

    dataview_endian_flip_changes_int16_value => {
        r#"const b=new ArrayBuffer(2); const v=new DataView(b); v.setInt16(0, 1, true); console.log(v.getInt16(0, false));"#,
        ["256"]
    };

    dataview_get_uint16_max_value => {
        r#"const b=new ArrayBuffer(2); const v=new DataView(b); v.setUint16(0, 0xffff, true); console.log(v.getUint16(0, true));"#,
        ["65535"]
    };

    dataview_get_int32_negative_little_endian => {
        r#"const b=new ArrayBuffer(4); const v=new DataView(b); v.setInt32(0, -1, true); console.log(v.getInt32(0, true));"#,
        ["-1"]
    };

    dataview_get_uint32_little_endian => {
        r#"const b=new ArrayBuffer(4); const v=new DataView(b); v.setUint32(0, 0x11223344, true); console.log(v.getUint32(0, true).toString(16));"#,
        ["11223344"]
    };

    dataview_get_float32_pi_approx => {
        r#"const b=new ArrayBuffer(4); const v=new DataView(b); v.setFloat32(0, 3.14, true); console.log(Math.round(v.getFloat32(0, true)*100));"#,
        ["314"]
    };

    dataview_get_float64_pi_precise => {
        r#"const b=new ArrayBuffer(8); const v=new DataView(b); v.setFloat64(0, Math.PI, true); console.log(v.getFloat64(0, true)>3.14);"#,
        ["true"]
    };

    dataview_get_bigint64_negative => {
        r#"const b=new ArrayBuffer(8); const v=new DataView(b); v.setBigInt64(0, -42n, true); console.log(v.getBigInt64(0, true));"#,
        ["-42"]
    };

    dataview_get_biguint64_large => {
        r#"const b=new ArrayBuffer(8); const v=new DataView(b); v.setBigUint64(0, 18446744073709551615n, true); console.log(v.getBigUint64(0, true));"#,
        ["18446744073709551615"]
    };

    dataview_byte_length_matches_buffer => {
        r#"const v=new DataView(new ArrayBuffer(16)); console.log(v.byteLength);"#,
        ["16"]
    };

    dataview_byte_offset_from_slice => {
        r#"const v=new DataView(new ArrayBuffer(8), 2, 4); console.log(v.byteOffset);console.log(v.byteLength);"#,
        ["2", "4"]
    };

    dataview_get_out_of_range_throws_range_error => {
        r#"const v=new DataView(new ArrayBuffer(1)); try{v.getInt16(0);}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    dataview_set_out_of_range_throws_range_error => {
        r#"const v=new DataView(new ArrayBuffer(1)); try{v.setInt32(0,1);}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    dataview_negative_index_throws_range_error => {
        r#"const v=new DataView(new ArrayBuffer(4)); try{v.getInt32(-1);}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    dataview_buffer_property_is_array_buffer => {
        r#"const buf=new ArrayBuffer(4); const v=new DataView(buf); console.log(v.buffer===buf);"#,
        ["true"]
    };

    dataview_set_int8_then_read_uint8_same_slot => {
        r#"const v=new DataView(new ArrayBuffer(1)); v.setInt8(0,-1); console.log(v.getUint8(0));"#,
        ["255"]
    };

    dataview_multiple_fields_in_one_buffer => {
        r#"const v=new DataView(new ArrayBuffer(6)); v.setUint16(0,1,true); v.setUint32(2,2,true); console.log(v.getUint16(0,true));console.log(v.getUint32(2,true));"#,
        ["1", "2"]
    };

    dataview_float32_nan_payload_preserved => {
        r#"const v=new DataView(new ArrayBuffer(4)); v.setFloat32(0, NaN, true); console.log(Number.isNaN(v.getFloat32(0,true)));"#,
        ["true"]
    };

    dataview_float64_negative_zero => {
        r#"const v=new DataView(new ArrayBuffer(8)); v.setFloat64(0,-0,true); console.log(1/v.getFloat64(0,true)<0);"#,
        ["true"]
    };

    dataview_int16_at_last_valid_offset => {
        r#"const v=new DataView(new ArrayBuffer(2)); v.setInt16(0,7,true); console.log(v.getInt16(0,true));"#,
        ["7"]
    };

    dataview_overlapping_int8_reads_from_int16_write => {
        r#"const v=new DataView(new ArrayBuffer(2)); v.setInt16(0,0x0102,true); console.log(v.getInt8(0));console.log(v.getInt8(1));"#,
        ["1", "2"]
    };

    dataview_instanceof_dataview => {
        r#"console.log(new DataView(new ArrayBuffer(1)) instanceof DataView);"#,
        ["true"]
    };

    dataview_from_typed_array_buffer_slice => {
        r#"const arr=new Uint8Array([1,2,3,4]); const v=new DataView(arr.buffer,1,2); console.log(v.getUint8(0));console.log(v.getUint8(1));"#,
        ["2", "3"]
    };

    dataview_set_float32_big_endian_differs_from_little => {
        r#"const v=new DataView(new ArrayBuffer(4)); v.setFloat32(0,1,true); const le=v.getFloat32(0,true); v.setFloat32(0,1,false); const be=v.getFloat32(0,false); console.log(le!==be);"#,
        ["true"]
    };

    dataview_get_int32_on_shared_buffer => {
        r#"const sab=new SharedArrayBuffer(4); const v=new DataView(sab); v.setInt32(0,99,true); console.log(v.getInt32(0,true));"#,
        ["99"]
    };
}
