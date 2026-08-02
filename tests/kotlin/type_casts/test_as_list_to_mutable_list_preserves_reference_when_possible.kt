// vybe-test: kotlin/type_casts/test_as_list_to_mutable_list_preserves_reference_when_possible
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val readonly: List<Int> = listOf(1, 2, 3)
            try {
                val mutable = readonly as MutableList<Int>
                mutable.add(4)
                println("mutated")
            } catch (e: Exception) {
                println("rejected")
            }
        }

