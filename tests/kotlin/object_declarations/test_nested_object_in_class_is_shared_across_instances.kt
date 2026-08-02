// vybe-test: kotlin/object_declarations/test_nested_object_in_class_is_shared_across_instances
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

class Holder {
            object Cache {
                var value = 0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Holder.Cache.value += 1
            Holder.Cache.value += 2
            __check((Holder.Cache.value).toString(), "3")
        }
