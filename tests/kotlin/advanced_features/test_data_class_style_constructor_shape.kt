// vybe-test: kotlin/advanced_features/test_data_class_style_constructor_shape
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

data class User(val name: String, val age: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = User("Ada", 25)
            __check((first.name).toString(), "Ada")
            __check((first.age).toString(), "25")
        }
