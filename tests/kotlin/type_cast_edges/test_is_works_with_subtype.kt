// vybe-test: kotlin/type_cast_edges/test_is_works_with_subtype
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

interface Animal { fun kind(): String }
        class Cat : Animal { override fun kind() = "cat" }
        class Dog : Animal { override fun kind() = "dog" }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Animal = Cat()
            __check((a is Animal).toString(), "true")
            __check((a is Cat).toString(), "true")
            __check((a is Dog).toString(), "false")
            __check((a.kind()).toString(), "cat")
        }
