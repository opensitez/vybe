//! hash/crc32, hash/adler32, hash/fnv runtime: Checksum, Update, New, Write, Sum —
//! distinct compile smokes in `test_crypto_hash_compile.rs` and `test_cover_hash_heap_io.rs`.


go_run_cases! {
    crc32_checksum_ieee_empty => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { fmt.Println(crc32.ChecksumIEEE([]byte{})) }",
        vec!["0"]
    ),
    crc32_checksum_ieee_single_byte => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { fmt.Println(crc32.ChecksumIEEE([]byte(\"a\"))) }",
        vec!["3904355907"]
    ),
    crc32_checksum_ieee_go_string => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { fmt.Println(crc32.ChecksumIEEE([]byte(\"go\"))) }",
        vec!["3060306774"]
    ),
    crc32_checksum_ieee_standard_vector => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { fmt.Println(crc32.ChecksumIEEE([]byte(\"123456789\"))) }",
        vec!["3421780262"]
    ),
    crc32_new_ieee_write_sum => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { h := crc32.NewIEEE(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["3060306774"]
    ),
    crc32_update_matches_checksum => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { table := crc32.IEEETable; direct := crc32.ChecksumIEEE([]byte(\"data\")); updated := crc32.Update(0, table, []byte(\"data\")); fmt.Println(direct == updated) }",
        vec!["true"]
    ),
    crc32_update_incremental => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { table := crc32.IEEETable; c := uint32(0); c = crc32.Update(c, table, []byte(\"ab\")); c = crc32.Update(c, table, []byte(\"c\")); full := crc32.ChecksumIEEE([]byte(\"abc\")); fmt.Println(c == full) }",
        vec!["true"]
    ),
    crc32_new_ieee_empty_write => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { h := crc32.NewIEEE(); fmt.Println(h.Sum32()) }",
        vec!["0"]
    ),
    crc32_make_table_castagnoli => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { t := crc32.MakeTable(crc32.Castagnoli); fmt.Println(t != nil) }",
        vec!["true"]
    ),
    crc32_checksum_castagnoli_differs_from_ieee => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { data := []byte(\"go\"); ieee := crc32.ChecksumIEEE(data); c := crc32.Checksum(data, crc32.MakeTable(crc32.Castagnoli)); fmt.Println(ieee != c) }",
        vec!["true"]
    ),
    adler32_checksum_empty => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { fmt.Println(adler32.Checksum([]byte{})) }",
        vec!["1"]
    ),
    adler32_checksum_go => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { fmt.Println(adler32.Checksum([]byte(\"go\"))) }",
        vec!["20906199"]
    ),
    adler32_checksum_wikipedia => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { fmt.Println(adler32.Checksum([]byte(\"Wikipedia\"))) }",
        vec!["300286872"]
    ),
    adler32_new_write_sum => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { h := adler32.New(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["20906199"]
    ),
    adler32_new_empty_sum => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { h := adler32.New(); fmt.Println(h.Sum32()) }",
        vec!["1"]
    ),
    adler32_incremental_write => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { h := adler32.New(); h.Write([]byte(\"g\")); h.Write([]byte(\"o\")); fmt.Println(h.Sum32()) }",
        vec!["20906199"]
    ),
    fnv_new32_empty_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New32(); fmt.Println(h.Sum32()) }",
        vec!["2166136261"]
    ),
    fnv_new32_write_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New32(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["1786192775"]
    ),
    fnv_new32a_write_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New32a(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["1109423947"]
    ),
    fnv_new64_empty_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New64(); fmt.Println(h.Sum64()) }",
        vec!["14695981039346656037"]
    ),
    fnv_new64_write_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New64(); h.Write([]byte(\"go\")); fmt.Println(h.Sum64()) }",
        vec!["590641186866933191"]
    ),
    fnv_new64a_write_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New64a(); h.Write([]byte(\"go\")); fmt.Println(h.Sum64()) }",
        vec!["618463229101696779"]
    ),
    crc32_differs_from_adler32_same_input => (
        "package main; import \"fmt\"; import \"hash/crc32\"; import \"hash/adler32\"; func main() { data := []byte(\"test\"); c := crc32.ChecksumIEEE(data); a := adler32.Checksum(data); fmt.Println(c != uint32(a)) }",
        vec!["true"]
    ),
    crc32_differs_from_fnv32_same_input => (
        "package main; import \"fmt\"; import \"hash/crc32\"; import \"hash/fnv\"; func main() { data := []byte(\"test\"); c := crc32.ChecksumIEEE(data); h := fnv.New32(); h.Write(data); fmt.Println(c != h.Sum32()) }",
        vec!["true"]
    ),
    adler32_differs_from_fnv32_same_input => (
        "package main; import \"fmt\"; import \"hash/adler32\"; import \"hash/fnv\"; func main() { data := []byte(\"test\"); a := adler32.Checksum(data); h := fnv.New32a(); h.Write(data); fmt.Println(uint32(a) != h.Sum32()) }",
        vec!["true"]
    ),
    fnv32_differs_from_fnv64_same_input => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { data := []byte(\"go\"); h32 := fnv.New32(); h32.Write(data); h64 := fnv.New64(); h64.Write(data); fmt.Println(h32.Sum32() != uint32(h64.Sum64())) }",
        vec!["true"]
    ),
    fnv32_differs_from_fnv32a_same_input => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { data := []byte(\"abc\"); h1 := fnv.New32(); h1.Write(data); h2 := fnv.New32a(); h2.Write(data); fmt.Println(h1.Sum32() != h2.Sum32()) }",
        vec!["true"]
    ),
    crc32_size_constant => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { fmt.Println(crc32.Size) }",
        vec!["4"]
    ),
    crc32_new_ieee_size_method => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { h := crc32.NewIEEE(); fmt.Println(h.Size()) }",
        vec!["4"]
    ),
    crc32_new_ieee_block_size => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { h := crc32.NewIEEE(); fmt.Println(h.BlockSize()) }",
        vec!["1"]
    ),
    adler32_size_via_hash_interface => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { h := adler32.New(); fmt.Println(h.Size()) }",
        vec!["4"]
    ),
    fnv_new32_block_size => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New32(); fmt.Println(h.BlockSize()) }",
        vec!["1"]
    ),
    crc32_sum_append_nil => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { h := crc32.NewIEEE(); h.Write([]byte(\"x\")); sum := h.Sum(nil); fmt.Println(len(sum)) }",
        vec!["4"]
    ),
    fnv_sum_append_nil => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New32(); h.Write([]byte(\"x\")); sum := h.Sum(nil); fmt.Println(len(sum)) }",
        vec!["4"]
    ),
    crc32_reset_then_rehash => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { h := crc32.NewIEEE(); h.Write([]byte(\"go\")); h.Reset(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["3060306774"]
    ),
    adler32_reset_then_rehash => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { h := adler32.New(); h.Write([]byte(\"go\")); h.Reset(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["20906199"]
    ),
    fnv_reset_then_rehash => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New64(); h.Write([]byte(\"go\")); h.Reset(); h.Write([]byte(\"go\")); fmt.Println(h.Sum64()) }",
        vec!["590641186866933191"]
    ),
    crc32_different_inputs_different_sums => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { a := crc32.ChecksumIEEE([]byte(\"a\")); b := crc32.ChecksumIEEE([]byte(\"b\")); fmt.Println(a != b) }",
        vec!["true"]
    ),
    adler32_different_inputs_different_sums => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { a := adler32.Checksum([]byte(\"a\")); b := adler32.Checksum([]byte(\"b\")); fmt.Println(a != b) }",
        vec!["true"]
    ),
    fnv_write_byte_slice_vs_string => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h1 := fnv.New32a(); h1.Write([]byte(\"go\")); h2 := fnv.New32a(); h2.Write([]byte(\"go\")); fmt.Println(h1.Sum32() == h2.Sum32()) }",
        vec!["true"]
    ),
    crc32_koopman_table_non_nil => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { t := crc32.MakeTable(crc32.Koopman); fmt.Println(t != nil) }",
        vec!["true"]
    ),
    fnv_new128_write_sum32_low => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New128(); h.Write([]byte(\"go\")); s := h.Sum(nil); fmt.Println(len(s)) }",
        vec!["16"]
    ),
    fnv_new128a_write_sum => (
        "package main; import \"fmt\"; import \"hash/fnv\"; func main() { h := fnv.New128a(); h.Write([]byte(\"go\")); s := h.Sum(nil); fmt.Println(len(s)) }",
        vec!["16"]
    ),
    crc32_hash_interface_sum32 => (
        "package main; import \"fmt\"; import \"hash/crc32\"; import \"hash\"; func main() { var h hash.Hash32 = crc32.NewIEEE(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["3060306774"]
    ),
    adler32_hash_interface_sum32 => (
        "package main; import \"fmt\"; import \"hash/adler32\"; import \"hash\"; func main() { var h hash.Hash32 = adler32.New(); h.Write([]byte(\"go\")); fmt.Println(h.Sum32()) }",
        vec!["20906199"]
    ),
    fnv_hash64_interface => (
        "package main; import \"fmt\"; import \"hash/fnv\"; import \"hash\"; func main() { var h hash.Hash64 = fnv.New64(); h.Write([]byte(\"go\")); fmt.Println(h.Sum64()) }",
        vec!["590641186866933191"]
    ),
    crc32_long_input_nonzero => (
        "package main; import \"fmt\"; import \"hash/crc32\"; func main() { data := make([]byte, 256); for i := range data { data[i] = byte(i) }; fmt.Println(crc32.ChecksumIEEE(data) != 0) }",
        vec!["true"]
    ),
    adler32_long_input_nonzero => (
        "package main; import \"fmt\"; import \"hash/adler32\"; func main() { data := make([]byte, 256); for i := range data { data[i] = byte(i) }; fmt.Println(adler32.Checksum(data) != 1) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    crc32_const_ieee => "package main; import \"hash/crc32\"; func main() { _ = crc32.IEEE }",
    crc32_const_castagnoli => "package main; import \"hash/crc32\"; func main() { _ = crc32.Castagnoli }",
    crc32_const_koopman => "package main; import \"hash/crc32\"; func main() { _ = crc32.Koopman }",
    fnv_new128a_empty => "package main; import \"hash/fnv\"; func main() { _ = fnv.New128a().Sum(nil) }",
    fnv_new64a_empty => "package main; import \"hash/fnv\"; func main() { _ = fnv.New64a().Sum(nil) }",
    crc32_new_castagnoli => "package main; import \"hash/crc32\"; func main() { h := crc32.New(crc32.MakeTable(crc32.Castagnoli)); _, _ = h.Write([]byte(\"x\")) }",
    adler32_write_nil_slice => "package main; import \"hash/adler32\"; func main() { h := adler32.New(); _, _ = h.Write(nil) }",
    fnv_write_empty_slice => "package main; import \"hash/fnv\"; func main() { h := fnv.New32(); _, _ = h.Write([]byte{}) }",
}
