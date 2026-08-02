// vybe-test: kotlin/data_class_copying/test_data_class_copy_zero_changes_is_identity_like
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Tag(val name: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Tag("x")
            val b = a.copy()
            __check((a).toString(), "Tag(name=x)")
            __check((b).toString(), "Tag(name=x)")
            __check((a == b).toString(), "true")
            __check((a === b).toString(), "false")
        }
