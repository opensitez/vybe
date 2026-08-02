// vybe-test: kotlin/local_classes/test_local_class_object_expression
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val o = object {
                val v = 4
                fun text() = "v${'$'}{v}"
            }
            __check((o.text()).toString(), "v4")
        }
