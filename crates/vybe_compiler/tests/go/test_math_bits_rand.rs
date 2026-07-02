//! math/bits (OnesCount, LeadingZeros) and math/rand (Intn, Seed).


go_run_cases! {
    bits_ones_count_zero => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(0)) }", vec!["0"]),
    bits_ones_count_one => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(1)) }", vec!["1"]),
    bits_ones_count_three => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(3)) }", vec!["2"]),
    bits_ones_count_fifteen => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(15)) }", vec!["4"]),
    bits_ones_count_all_ones_byte => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(255)) }", vec!["8"]),
    bits_ones_count_power_of_two => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(16)) }", vec!["1"]),
    bits_ones_count_alternating_bits => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(0xAAAAAAAA)) }", vec!["16"]),
    bits_ones_count_negative_one => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(-1)) }", vec!["64"]),
    bits_ones_count_max_int => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(^0)) }", vec!["64"]),
    bits_ones_count_sparse_high_bit => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.OnesCount(1 << 31)) }", vec!["1"]),
    bits_leading_zeros_zero => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(0)) }", vec!["64"]),
    bits_leading_zeros_one => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(1)) }", vec!["63"]),
    bits_leading_zeros_eight => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(8)) }", vec!["60"]),
    bits_leading_zeros_high_bit_set => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(1 << 63)) }", vec!["0"]),
    bits_leading_zeros_max_int => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(^0)) }", vec!["0"]),
    bits_leading_zeros_negative_one => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(-1)) }", vec!["0"]),
    bits_leading_zeros_power_of_two => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(1024)) }", vec!["53"]),
    bits_leading_zeros_all_but_low_bit => ("package main; import \"fmt\"; import \"math/bits\"; func main() { fmt.Println(bits.LeadingZeros(2)) }", vec!["62"]),
    rand_intn_non_negative => ("package main; import \"fmt\"; import \"math/rand\"; func main() { fmt.Println(rand.Intn(10) >= 0) }", vec!["true"]),
    rand_intn_strictly_less_than_bound => ("package main; import \"fmt\"; import \"math/rand\"; func main() { n := rand.Intn(10); fmt.Println(n >= 0 && n < 10) }", vec!["true"]),
    rand_intn_unit_bound_non_negative => ("package main; import \"fmt\"; import \"math/rand\"; func main() { fmt.Println(rand.Intn(1) >= 0) }", vec!["true"]),
    rand_intn_large_bound_in_range => ("package main; import \"fmt\"; import \"math/rand\"; func main() { fmt.Println(rand.Intn(1000000) < 1000000) }", vec!["true"]),
    rand_intn_two_draws_bounded => ("package main; import \"fmt\"; import \"math/rand\"; func main() { a := rand.Intn(5); b := rand.Intn(5); fmt.Println(a >= 0 && b >= 0 && a < 5 && b < 5) }", vec!["true"]),
    rand_intn_loop_accumulates_valid => ("package main; import \"fmt\"; import \"math/rand\"; func main() { ok := 0; i := 0; for i < 4 { if rand.Intn(3) < 3 { ok++ }; i++ }; fmt.Println(ok) }", vec!["4"]),
}

go_compile_cases! {
    bits_ones_count_uint8 => "package main; import \"math/bits\"; func main() { _ = bits.OnesCount8(0xFF) }",
    bits_ones_count_uint16 => "package main; import \"math/bits\"; func main() { _ = bits.OnesCount16(0xFFFF) }",
    bits_ones_count_uint32 => "package main; import \"math/bits\"; func main() { _ = bits.OnesCount32(0xFFFFFFFF) }",
    bits_ones_count_uint64 => "package main; import \"math/bits\"; func main() { _ = bits.OnesCount64(^uint64(0)) }",
    bits_leading_zeros_uint8 => "package main; import \"math/bits\"; func main() { _ = bits.LeadingZeros8(1) }",
    bits_leading_zeros_uint16 => "package main; import \"math/bits\"; func main() { _ = bits.LeadingZeros16(1) }",
    bits_leading_zeros_uint32 => "package main; import \"math/bits\"; func main() { _ = bits.LeadingZeros32(1) }",
    bits_leading_zeros_uint64 => "package main; import \"math/bits\"; func main() { _ = bits.LeadingZeros64(1) }",
    rand_seed_then_intn => "package main; import \"math/rand\"; func main() { rand.Seed(42); _ = rand.Intn(100) }",
    rand_seed_reproducible_sequence => "package main; import \"fmt\"; import \"math/rand\"; func main() { rand.Seed(1); fmt.Println(rand.Intn(10)); rand.Seed(1); fmt.Println(rand.Intn(10)) }",
    rand_seed_zero => "package main; import \"math/rand\"; func main() { rand.Seed(0); _ = rand.Intn(5) }",
    rand_seed_before_multiple_intn => "package main; import \"math/rand\"; func main() { rand.Seed(99); _ = rand.Intn(10); _ = rand.Intn(10) }",
}
