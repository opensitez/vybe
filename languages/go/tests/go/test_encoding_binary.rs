//! encoding/binary: BigEndian, LittleEndian, Put/Uint, varint, Size, Read, Write.

go_run_cases! {
    binary_big_endian_put_uint16_byte_layout => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 2); binary.BigEndian.PutUint16(buf, 0x0102); fmt.Println(int(buf[0])); fmt.Println(int(buf[1])) }",
        vec!["1", "2"]
    ),
    binary_little_endian_put_uint16_byte_layout => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 2); binary.LittleEndian.PutUint16(buf, 0x0102); fmt.Println(int(buf[0])); fmt.Println(int(buf[1])) }",
        vec!["2", "1"]
    ),
    binary_big_endian_uint16_from_two_bytes => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { fmt.Println(binary.BigEndian.Uint16([]byte{0x01, 0x02})) }",
        vec!["258"]
    ),
    binary_little_endian_uint16_from_two_bytes => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { fmt.Println(binary.LittleEndian.Uint16([]byte{0x02, 0x01})) }",
        vec!["258"]
    ),
    binary_big_endian_put_uint32_byte_layout => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 4); binary.BigEndian.PutUint32(buf, 0x01020304); fmt.Println(int(buf[0])); fmt.Println(int(buf[3])) }",
        vec!["1", "4"]
    ),
    binary_little_endian_put_uint32_byte_layout => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 4); binary.LittleEndian.PutUint32(buf, 0x01020304); fmt.Println(int(buf[0])); fmt.Println(int(buf[3])) }",
        vec!["4", "1"]
    ),
    binary_big_endian_uint32_from_four_bytes => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { fmt.Println(binary.BigEndian.Uint32([]byte{0, 0, 0, 0x2a})) }",
        vec!["42"]
    ),
    binary_little_endian_uint32_from_four_bytes => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { fmt.Println(binary.LittleEndian.Uint32([]byte{0x2a, 0, 0, 0})) }",
        vec!["42"]
    ),
    binary_big_endian_put_uint64_first_and_last_byte => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 8); binary.BigEndian.PutUint64(buf, 0x0102030405060708); fmt.Println(int(buf[0])); fmt.Println(int(buf[7])) }",
        vec!["1", "8"]
    ),
    binary_little_endian_put_uint64_low_byte_first => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 8); binary.LittleEndian.PutUint64(buf, 0x0102030405060708); fmt.Println(int(buf[0])); fmt.Println(int(buf[7])) }",
        vec!["8", "1"]
    ),
    binary_big_endian_put_int16_negative_one => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 2); binary.BigEndian.PutInt16(buf, -1); fmt.Println(int(buf[0])); fmt.Println(int(buf[1])) }",
        vec!["255", "255"]
    ),
    binary_little_endian_put_int16_negative_one => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, 2); binary.LittleEndian.PutInt16(buf, -1); fmt.Println(int(buf[0])); fmt.Println(int(buf[1])) }",
        vec!["255", "255"]
    ),
    binary_big_endian_int32_from_slice => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { fmt.Println(binary.BigEndian.Int32([]byte{0xff, 0xff, 0xff, 0xff})) }",
        vec!["-1"]
    ),
    binary_put_uvarint_single_byte_max => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, binary.MaxVarintLen64); n := binary.PutUvarint(buf, 127); fmt.Println(n); fmt.Println(int(buf[0])) }",
        vec!["1", "127"]
    ),
    binary_put_uvarint_two_byte_encoding_for_128 => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, binary.MaxVarintLen64); n := binary.PutUvarint(buf, 128); fmt.Println(n); fmt.Println(int(buf[0])); fmt.Println(int(buf[1])) }",
        vec!["2", "128", "1"]
    ),
    binary_uvarint_reads_multi_byte_value => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { v, n := binary.Uvarint([]byte{0x80, 0x01}); fmt.Println(v); fmt.Println(n) }",
        vec!["128", "2"]
    ),
    binary_put_varint_negative_one_single_byte => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { buf := make([]byte, binary.MaxVarintLen64); n := binary.PutVarint(buf, -1); fmt.Println(n); fmt.Println(int(buf[0])) }",
        vec!["1", "1"]
    ),
    binary_varint_reads_negative_one => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { v, n := binary.Varint([]byte{0x01}); fmt.Println(v); fmt.Println(n) }",
        vec!["-1", "1"]
    ),
    binary_size_uint16_field_is_two => (
        "package main; import \"fmt\"; import \"encoding/binary\"; func main() { fmt.Println(binary.Size(uint16(0))) }",
        vec!["2"]
    ),
    binary_read_uint16_from_bytes_reader => (
        "package main; import \"bytes\"; import \"fmt\"; import \"encoding/binary\"; func main() { r := bytes.NewReader([]byte{0x01, 0x02}); var v uint16; _ = binary.Read(r, binary.BigEndian, &v); fmt.Println(v) }",
        vec!["258"]
    ),
}

go_compile_cases! {
    binary_append_uint16_extends_empty_slice => "package main; import \"encoding/binary\"; func main() { b := []byte{}; _ = binary.BigEndian.AppendUint16(b, 0x0102) }",
    binary_append_uint32_to_slice => "package main; import \"encoding/binary\"; func main() { b := make([]byte, 0, 8); _ = binary.LittleEndian.AppendUint32(b, 42) }",
    binary_append_uvarint_to_slice => "package main; import \"encoding/binary\"; func main() { b := []byte{}; _ = binary.AppendUvarint(b, 300) }",
    binary_read_decodes_struct_fields => "package main; import \"bytes\"; import \"encoding/binary\"; type Header struct { Magic uint16; Ver uint8 }; func main() { var h Header; _ = binary.Read(bytes.NewReader([]byte{0xbe, 0xef, 0x01}), binary.BigEndian, &h) }",
    binary_write_uint16_via_new_buffer => "package main; import \"bytes\"; import \"encoding/binary\"; func main() { buf := bytes.NewBuffer(nil); _ = binary.Write(buf, binary.BigEndian, uint16(0x0102)) }",
    binary_read_full_exact_byte_count => "package main; import \"bytes\"; import \"encoding/binary\"; func main() { r := bytes.NewReader([]byte{1, 2, 3, 4}); dst := make([]byte, 4); _, _ = binary.ReadFull(r, dst) }",
    binary_native_endian_put_uint16 => "package main; import \"encoding/binary\"; func main() { buf := make([]byte, 2); binary.NativeEndian.PutUint16(buf, 0xabcd) }",
    binary_big_endian_put_float32 => "package main; import \"encoding/binary\"; import \"math\"; func main() { buf := make([]byte, 4); binary.BigEndian.PutUint32(buf, math.Float32bits(1.5)) }",
    binary_little_endian_put_float64 => "package main; import \"encoding/binary\"; import \"math\"; func main() { buf := make([]byte, 8); binary.LittleEndian.PutUint64(buf, math.Float64bits(2.5)) }",
    binary_big_endian_float32_from_bits => "package main; import \"encoding/binary\"; import \"math\"; func main() { _ = math.Float32frombits(binary.BigEndian.Uint32([]byte{0x3f, 0xc0, 0, 0})) }",
}
