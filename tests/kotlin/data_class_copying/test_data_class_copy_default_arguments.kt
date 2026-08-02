// vybe-test: kotlin/data_class_copying/test_data_class_copy_default_arguments
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Person(val name: String, val age: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Person("a", 1)
            val b = a.copy(age = 2)
            val c = a.copy(name = "b")
            __check((b.name).toString(), "a")
            __check((b.age).toString(), "2")
            __check((c.name).toString(), "b")
            __check((c.age).toString(), "1")
        }
