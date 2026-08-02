// vybe-test: kotlin/when_subjects/test_when_subject_object_type
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

open class Animal
        class Dog : Animal()
        class Cat : Animal()
        fun identify(a: Animal): String = when (a) {
            is Dog -> "dog"
            is Cat -> "cat"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((identify(Dog())).toString(), "dog")
            __check((identify(Cat())).toString(), "cat")
        }
