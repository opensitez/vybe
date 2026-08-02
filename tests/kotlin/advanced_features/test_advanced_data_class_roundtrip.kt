// vybe-test: kotlin/advanced_features/test_advanced_data_class_roundtrip
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

data class User(val name: String, val age: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = User("ivy", 11)
            val second = User("ivy", 11)
            __check((first == second).toString(), "true")
            __check((first.name).toString(), "ivy")
            __check((first.age).toString(), "11")
        }
