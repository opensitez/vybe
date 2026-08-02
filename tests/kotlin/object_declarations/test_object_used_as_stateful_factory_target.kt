// vybe-test: kotlin/object_declarations/test_object_used_as_stateful_factory_target
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Factory {
            fun create(label: String): Holder = Holder(label)
        }

        class Holder(val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Factory.create("x").label).toString(), "x")
        }
