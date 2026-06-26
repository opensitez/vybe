//! crypto/sha256 and hash interfaces — compile coverage for hashing stdlib.

use crate::helpers::*;

go_run_cases! {
    sha256_sum_empty => ("package main; import \"fmt\"; import \"crypto/sha256\"; func main() { sum := sha256.Sum256([]byte{}); fmt.Println(len(sum)) }", vec!["32"]),
    sha256_sum_abc => ("package main; import \"fmt\"; import \"crypto/sha256\"; func main() { sum := sha256.Sum256([]byte(\"abc\")); fmt.Println(int(sum[0])) }", vec!["186"]),
}

go_compile_cases! {
    sha256_new_writer => "package main; import \"crypto/sha256\"; func main() { h := sha256.New(); _, _ = h.Write([]byte(\"data\")) }",
    sha256_size_constant => "package main; import \"crypto/sha256\"; func main() { _ = sha256.Size }",
    md5_sum_compile => "package main; import \"crypto/md5\"; func main() { _ = md5.Sum([]byte(\"x\")) }",
    sha1_sum_compile => "package main; import \"crypto/sha1\"; func main() { _ = sha1.Sum([]byte(\"x\")) }",
    hash_crc32_compile => "package main; import \"hash/crc32\"; func main() { _ = crc32.ChecksumIEEE([]byte(\"go\")) }",
    hash_fnv_compile => "package main; import \"hash/fnv\"; func main() { h := fnv.New32a(); _, _ = h.Write([]byte(\"a\")) }",
}
