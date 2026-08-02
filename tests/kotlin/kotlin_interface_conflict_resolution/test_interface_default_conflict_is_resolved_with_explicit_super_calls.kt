// vybe-test: kotlin/kotlin_interface_conflict_resolution/test_interface_default_conflict_is_resolved_with_explicit_super_calls
// origin: languages/kotlin/tests/kotlin/test_kotlin_interface_conflict_resolution.rs

interface First {
            fun origin(): String = "first"
        }

        interface Second {
            fun origin(): String = "second"
        }

        class Composite : First, Second {
            override fun origin(): String = super<First>.origin() + "/" + super<Second>.origin()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Composite().origin()).toString(), "first/second")
        }
