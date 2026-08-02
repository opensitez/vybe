// vybe-test: kotlin/object_expressions/test_object_expression_implements_multiple_methods
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface A {
            fun a(): Int
        }

        interface B {
            fun b(): Int
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val combined = object : A, B {
                override fun a(): Int { return 1 }
                override fun b(): Int { return 2 }
            }
            __check((combined.a() + combined.b()).toString(), "3")
        }
