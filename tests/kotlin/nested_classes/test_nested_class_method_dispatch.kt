// vybe-test: kotlin/nested_classes/test_nested_class_method_dispatch
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Service {
            class State(val ok: Boolean)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Service.State(true)
            val b = Service.State(false)
            val total = (if (a.ok) 1 else 0) + (if (b.ok) 1 else 0)
            __check((total).toString(), "1")
        }
