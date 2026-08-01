use crate::helpers::run_prints;

#[test]
fn test_companion_object_holds_shared_state() {
    let out = run_prints(
        r#"
        class Sequence {
            private val id: Int

            private constructor(value: Int) {
                id = value
            }

            companion object {
                var next: Int = 0
                fun nextSequence(): Sequence {
                    next += 1
                    return Sequence(next)
                }
            }

            fun value(): Int = id
        }

        fun main() {
            val a = Sequence.nextSequence()
            val b = Sequence.nextSequence()
            println(a.value())
            println(b.value())
            println(Sequence.next)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "2"]);
}

#[test]
fn test_companion_access_as_property_like_accessor() {
    let out = run_prints(
        r#"
        class Counter {
            companion object {
                var current: Int = 0

                fun bump(): Int {
                    current += 1
                    return current
                }
            }
        }

        fun main() {
            println(Counter.bump())
            println(Counter.current)
        }
    "#,
    );
    assert_eq!(out, &["1", "1"]);
}
