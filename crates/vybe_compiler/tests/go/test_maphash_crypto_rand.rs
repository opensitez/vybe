//! hash/maphash and crypto/rand compile coverage.

use crate::helpers::*;

go_compile_cases! {
    maphash_new => "package main; import \"hash/maphash\"; func main() { var h maphash.Hash; h.SetSeed(maphash.MakeSeed()) }",
    maphash_write_string => "package main; import \"hash/maphash\"; func main() { var h maphash.Hash; h.WriteString(\"go\") }",
    maphash_write_byte => "package main; import \"hash/maphash\"; func main() { var h maphash.Hash; h.WriteByte('x') }",
    maphash_sum64 => "package main; import \"hash/maphash\"; func main() { var h maphash.Hash; _ = h.Sum64() }",
    maphash_reset => "package main; import \"hash/maphash\"; func main() { var h maphash.Hash; h.Reset() }",
    maphash_bytes => "package main; import \"hash/maphash\"; func main() { var h maphash.Hash; _ = h.Bytes() }",
    rand_read_compile => "package main; import \"crypto/rand\"; func main() { b := make([]byte, 8); _, _ = rand.Read(b) }",
    rand_int_compile => "package main; import \"crypto/rand\"; import \"math/big\"; func main() { _, _ = rand.Int(rand.Reader, big.NewInt(10)) }",
    rand_prime_compile => "package main; import \"crypto/rand\"; import \"math/big\"; func main() { _, _ = rand.Prime(rand.Reader, 16) }",
}

go_run_cases! {
    maphash_string_sum => ("package main; import \"fmt\"; import \"hash/maphash\"; func main() { var h maphash.Hash; h.SetSeed(maphash.MakeSeed()); h.WriteString(\"vybe\"); fmt.Println(h.Sum64() > 0) }", vec!["true"]),
}
