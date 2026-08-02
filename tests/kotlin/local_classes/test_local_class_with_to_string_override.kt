// vybe-test: kotlin/local_classes/test_local_class_with_to_string_override
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class L(val v: Int) {
                override fun toString() = "val=" + v
            }
            __check((L(8).toString()).toString(), "val=8")
        }
