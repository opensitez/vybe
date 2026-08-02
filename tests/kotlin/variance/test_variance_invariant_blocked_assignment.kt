// vybe-test: kotlin/variance/test_variance_invariant_blocked_assignment
// origin: languages/kotlin/tests/kotlin/test_variance.rs

open class Animal
        class Dog : Animal()
        fun <T> accepts(list: MutableList<T>) {
            __check((list.size).toString(), "1")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val dogs: MutableList<Dog> = mutableListOf(Dog())
            accepts(dogs)
        }
