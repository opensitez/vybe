// vybe-test: kotlin/nested_classes/test_nested_class_local_type_identity
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Container {
            class Marker
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first: Container.Marker = Container.Marker()
            val second = Container.Marker()
            __check((first == second).toString(), "false")
            __check((first != null).toString(), "true")
        }
