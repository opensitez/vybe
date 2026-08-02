// vybe-test: kotlin/data_classes/test_data_class_hash_code_matches_equality
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Key(val id: Int, val tag: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Key(10, "tag")
            val b = Key(10, "tag")
            __check((a.hashCode() == b.hashCode()).toString(), "true")
            __check((a.hashCode() != a.hashCode()).toString(), "false")
        }
