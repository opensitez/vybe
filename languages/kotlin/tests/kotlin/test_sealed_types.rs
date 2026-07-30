use crate::helpers::run_prints;

#[test]
fn test_simple_sealed_when_exhaustive_without_else() {
    let out = run_prints(r#"
        sealed class Result {
            class Ok(val value: Int) : Result()
            class Fail : Result()
        }

        fun describe(result: Result): String {
            return when (result) {
                is Result.Ok -> "ok:" + result.value.toString()
                is Result.Fail -> "fail"
            }
        }

        fun main() {
            val value = describe(Result.Ok(3))
            val other = describe(Result.Fail())
            println(value)
            println(other)
        }
    "#);
    assert_eq!(out, &["ok:3", "fail"]);
}

#[test]
fn test_nested_sealed_hierarchy_with_data() {
    let out = run_prints(r#"
        sealed class Command {
            data class Print(val value: String) : Command()
            data class Count(val value: Int) : Command()
        }

        fun execute(command: Command): String {
            return when (command) {
                is Command.Print -> command.value
                is Command.Count -> "count=" + command.value.toString()
            }
        }

        fun main() {
            println(execute(Command.Print("x")))
            println(execute(Command.Count(4)))
        }
    "#);
    assert_eq!(out, &["x", "count=4"]);
}

#[test]
fn test_sealed_class_with_leaf_subclass_instances() {
    let out = run_prints(r#"
        sealed class Node {
            class Leaf(val value: Int) : Node()
            class Branch(val left: Node, val right: Node) : Node()
        }

        fun count(node: Node): Int {
            return when (node) {
                is Node.Leaf -> 1
                is Node.Branch -> 1 + count(node.left) + count(node.right)
            }
        }

        fun main() {
            val root = Node.Branch(Node.Leaf(1), Node.Leaf(2))
            println(count(root))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_sealed_interface_like_hierarchy() {
    let out = run_prints(r#"
        sealed interface Event

        class Start : Event
        class Stop : Event

        fun label(event: Event): String {
            return when (event) {
                is Start -> "start"
                is Stop -> "stop"
            }
        }

        fun main() {
            println(label(Start()))
            println(label(Stop()))
        }
    "#);
    assert_eq!(out, &["start", "stop"]);
}

#[test]
fn test_sealed_with_same_name_companion_members() {
    let out = run_prints(r#"
        sealed class State {
            class Active : State()
            class Error : State()

            companion object {
                fun active(): State = Active()
            }
        }

        fun main() {
            val state = State.active()
            println(when (state) {
                is State.Active -> 1
                is State.Error -> 0
            })
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_sealed_leafs_can_hold_state_in_constructor() {
    let out = run_prints(r#"
        sealed class Packet {
            class Text(val text: String) : Packet()
            class Number(val value: Int) : Packet()
        }

        fun render(packet: Packet): String {
            return when (packet) {
                is Packet.Text -> packet.text
                is Packet.Number -> "n=" + packet.value.toString()
            }
        }

        fun main() {
            println(render(Packet.Text("z")))
            println(render(Packet.Number(6)))
        }
    "#);
    assert_eq!(out, &["z", "n=6"]);
}

#[test]
fn test_deeply_nested_when_on_sealed_tree() {
    let out = run_prints(r#"
        sealed class Expr {
            class Value(val value: Int) : Expr()
            class Add(val left: Expr, val right: Expr) : Expr()
            class Negate(val source: Expr) : Expr()
        }

        fun evaluate(expr: Expr): Int {
            return when (expr) {
                is Expr.Value -> expr.value
                is Expr.Negate -> -evaluate(expr.source)
                is Expr.Add -> evaluate(expr.left) + evaluate(expr.right)
            }
        }

        fun main() {
            val expr = Expr.Add(Expr.Value(3), Expr.Negate(Expr.Value(2)))
            println(evaluate(expr))
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_non_exhaustive_when_still_requires_else_for_non_sealed() {
    let out = run_prints(r#"
        sealed class Alpha {
            class A : Alpha()
        }

        open class Beta

        fun main() {
            val alpha: Alpha = Alpha.A()
            println(when (alpha) {
                is Alpha.A -> 1
            })
            val beta = Beta()
            println(when (beta is Beta) {
                true -> 2
                false -> 3
            })
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_sealed_direct_instance_creation_from_subclass() {
    let out = run_prints(r#"
        sealed class Unit {
            class Meter(val value: Int) : Unit()
        }

        fun main() {
            val unit = Unit.Meter(5)
            println(unit.value)
        }
    "#);
    assert_eq!(out, &["5"]);
}

#[test]
fn test_sealed_hierarchy_with_multiple_branches() {
    let out = run_prints(r#"
        sealed class Token {
            class A : Token()
            class B : Token()
            class C : Token()
        }

        fun score(token: Token): Int {
            return when (token) {
                is Token.A -> 1
                is Token.B -> 2
                is Token.C -> 3
            }
        }

        fun main() {
            println(score(Token.A()))
            println(score(Token.B()))
            println(score(Token.C()))
        }
    "#);
    assert_eq!(out, &["1", "2", "3"]);
}
