// vybe-test: kotlin/variance/test_variance_out_projection_setter_not_allowed_compile_time_skipped
// origin: languages/kotlin/tests/kotlin/test_variance.rs

class Repo<T> {
            private val store = mutableListOf<T>()
            fun getStore(): List<T> = store
            fun addAll(values: List<out T>) {
                // this path is intentionally empty
            }
            fun mainAdd() {
                __check((store.size).toString(), "0")
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = Repo<String>()
            r.mainAdd()
        }
