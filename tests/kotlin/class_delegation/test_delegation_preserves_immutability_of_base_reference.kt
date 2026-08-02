// vybe-test: kotlin/class_delegation/test_delegation_preserves_immutability_of_base_reference
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface View { fun size(): Int }

        class Snapshot(private val items: List<Int>) : View {
            override fun size() = items.size
        }

        class SnapshotWrapper(delegate: View) : View by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = Snapshot(listOf(1, 2))
            val wrapped = SnapshotWrapper(original)
            __check((wrapped.size()).toString(), "2")
        }
