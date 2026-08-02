// vybe-test: kotlin/kotlin_type_parameter_bounds/test_type_parameter_with_enum_constraint
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

interface Named { fun name(): String }

        enum class Source : Named { A { override fun name() = "a" }, B { override fun name() = "b" } }

        fun <T> label(item: T): String where T : Enum<T>, T : Named {
            return item.name()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(Source.A)).toString(), "a")
            __check((label(Source.B)).toString(), "b")
        }
