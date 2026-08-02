// vybe-test: kotlin/object_declarations/test_object_with_private_state_exposes_only_public_api
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Registry {
            private var next = 0
            fun next(): Int {
                next += 1
                return next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Registry.next()).toString(), "1")
            __check((Registry.next()).toString(), "2")
        }
