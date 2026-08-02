// vybe-test: kotlin/generics/test_generic_recursive_data_structure_preserves_type
// origin: languages/kotlin/tests/kotlin/test_generics.rs

data class Node<T>(val value: T, val next: Node<T>? = null)

        fun <T> collect(values: Node<T>): String {
            var cursor: Node<T>? = values
            var out = ""
            while (cursor != null) {
                out += cursor.value.toString()
                cursor = cursor.next
                if (cursor != null) out += "-"
            }
            return out
        }

        fun main() {
            val chain = Node("a", Node("b", Node("c")))
            println(collect(chain))
        }

