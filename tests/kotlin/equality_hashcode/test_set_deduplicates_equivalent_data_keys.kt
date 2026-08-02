// vybe-test: kotlin/equality_hashcode/test_set_deduplicates_equivalent_data_keys
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Key(val id: String, val version: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = hashSetOf(Key("x", 1), Key("x", 1), Key("x", 2))
            __check((set.size).toString(), "2")
            __check((set.contains(Key("x", 2))).toString(), "true")
            __check((set.contains(Key("x", 3))).toString(), "false")
        }
