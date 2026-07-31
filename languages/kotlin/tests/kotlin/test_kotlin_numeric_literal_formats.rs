kotlin_run_cases! {
    test_hex_and_binary_literals => (r#"
        fun main() {
            val hex = 0x10
            val bin = 0b10
            val oct = 8
            println(hex)
            println(bin)
            println(oct)
        }
    "#, vec!["16", "2", "8"]),
    test_numeric_underscores => (r#"
        fun main() {
            val million = 1_000_000
            val grouped = 12_34_56
            println(million)
            println(grouped)
        }
    "#, vec!["1000000", "123456"]),
    test_long_and_float_forms => (r#"
        fun main() {
            val big = 1_000_000L
            val small = 1.5
            println(big)
            println(small.toString())
        }
    "#, vec!["1000000", "1.5"]),
}
