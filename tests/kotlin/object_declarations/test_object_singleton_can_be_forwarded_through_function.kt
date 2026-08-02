// vybe-test: kotlin/object_declarations/test_object_singleton_can_be_forwarded_through_function
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Service {
            fun id(): Int = 1
        }

        fun getService(): Any {
            return Service
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((getService() === Service).toString(), "true")
            __check((getService() !== null).toString(), "true")
            __check((Service.id()).toString(), "1")
        }
