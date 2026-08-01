use crate::helpers::run_prints;

#[test]
fn test_simple_sealed_when_exhaustive_without_else() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["ok:3", "fail"]);
}

#[test]
fn test_nested_sealed_hierarchy_with_data() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["x", "count=4"]);
}

#[test]
fn test_sealed_class_with_leaf_subclass_instances() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_sealed_interface_like_hierarchy() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["start", "stop"]);
}

#[test]
fn test_sealed_with_same_name_companion_members() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_sealed_leafs_can_hold_state_in_constructor() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["z", "n=6"]);
}

#[test]
fn test_deeply_nested_when_on_sealed_tree() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_non_exhaustive_when_still_requires_else_for_non_sealed() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_sealed_direct_instance_creation_from_subclass() {
    let out = run_prints(
        r#"
        sealed class Unit {
            class Meter(val value: Int) : Unit()
        }

        fun main() {
            val unit = Unit.Meter(5)
            println(unit.value)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_sealed_hierarchy_with_multiple_branches() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_sealed_local_class_still_enforces_local_exhaustiveness() {
    let out = run_prints(
        r#"
        fun decide(state: State): String {
            return when (state) {
                is State.Ok -> "ok"
                is State.Fail -> "fail"
                is State.Ignore -> "ignore"
            }
        }

        sealed class State {
            class Ok : State()
            class Fail : State()
            class Ignore : State()
        }

        fun main() {
            println(decide(State.Ok()))
            println(decide(State.Fail()))
            println(decide(State.Ignore()))
        }
    "#,
    );
    assert_eq!(out, &["ok", "fail", "ignore"]);
}

#[test]
fn test_sealed_class_with_object_leaf() {
    let out = run_prints(
        r#"
        sealed class Option {
            object Empty : Option()
            class Value(val value: Int) : Option()
        }

        fun label(value: Option): String {
            return when (value) {
                is Option.Empty -> "empty"
                is Option.Value -> value.value.toString()
            }
        }

        fun main() {
            println(label(Option.Empty))
            println(label(Option.Value(5)))
        }
    "#,
    );
    assert_eq!(out, &["empty", "5"]);
}

#[test]
fn test_sealed_interface_and_data_leaves() {
    let out = run_prints(
        r#"
        sealed interface Kind

        data class Node(val id: Int) : Kind
        class Done : Kind

        fun describe(kind: Kind): String {
            return when (kind) {
                is Node -> "node:" + kind.id.toString()
                is Done -> "done"
            }
        }

        fun main() {
            println(describe(Node(7)))
            println(describe(Done()))
        }
    "#,
    );
    assert_eq!(out, &["node:7", "done"]);
}

#[test]
fn test_nested_sealed_subclasses_keep_disjoint_branches() {
    let out = run_prints(
        r#"
        sealed class Shape {
            sealed class Circle : Shape() {
                class Small : Circle()
                class Large : Circle()
            }

            class Square : Shape()
        }

        fun area(shape: Shape): String {
            return when (shape) {
                is Shape.Circle.Small -> "small"
                is Shape.Circle.Large -> "large"
                is Shape.Square -> "square"
            }
        }

        fun main() {
            println(area(Shape.Circle.Small()))
            println(area(Shape.Circle.Large()))
            println(area(Shape.Square()))
        }
    "#,
    );
    assert_eq!(out, &["small", "large", "square"]);
}

#[test]
fn test_sealed_type_with_generic_payload() {
    let out = run_prints(
        r#"
        sealed class Result<T> {
            class Ok<T>(val value: T) : Result<T>()
            class Error<T>(val reason: String) : Result<T>()
        }

        fun render(value: Result<Int>): String {
            return when (value) {
                is Result.Ok -> "ok:" + value.value.toString()
                is Result.Error -> "err:" + value.reason
            }
        }

        fun main() {
            println(render(Result.Ok(4)))
            println(render(Result.Error("x")))
        }
    "#,
    );
    assert_eq!(out, &["ok:4", "err:x"]);
}

#[test]
fn test_when_on_sealed_class_in_function_scope() {
    let out = run_prints(
        r#"
        sealed class Command
        class Start : Command()
        class Stop : Command()

        fun describe(command: Command): String {
            return when (command) {
                is Start -> "start"
                is Stop -> "stop"
            }
        }

        fun main() {
            val command: Command = if (true) Start() else Stop()
            println(describe(command))
        }
    "#,
    );
    assert_eq!(out, &["start"]);
}

#[test]
fn test_sealed_dispatch_supports_recursive_visitation() {
    let out = run_prints(
        r#"
        sealed class Node {
            class Leaf(val value: Int) : Node()
            class Branch(val left: Node, val right: Node) : Node()
        }

        fun sum(node: Node): Int {
            return when (node) {
                is Node.Leaf -> node.value
                is Node.Branch -> sum(node.left) + sum(node.right)
            }
        }

        fun main() {
            val tree = Node.Branch(Node.Leaf(1), Node.Branch(Node.Leaf(2), Node.Leaf(3)))
            println(sum(tree))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_sealed_class_can_be_used_as_enum_like_protocol() {
    let out = run_prints(
        r#"
        sealed class Status {
            class Success(val message: String) : Status()
            class Failure(val code: Int) : Status()
        }

        fun statusCode(status: Status): Int {
            return when (status) {
                is Status.Success -> 0
                is Status.Failure -> status.code
            }
        }

        fun main() {
            println(statusCode(Status.Success("ok")))
            println(statusCode(Status.Failure(7)))
        }
    "#,
    );
    assert_eq!(out, &["0", "7"]);
}

#[test]
fn test_sealed_when_expression_with_else_keeps_runtime_branch() {
    let out = run_prints(
        r#"
        sealed class Token {
            class Left : Token()
            class Right : Token()
        }

        fun describe(token: Token): String {
            return when (token) {
                is Token.Left -> "left"
                is Token.Right -> "right"
                else -> "other"
            }
        }

        fun main() {
            println(describe(Token.Left()))
            println(describe(Token.Right()))
        }
    "#,
    );
    assert_eq!(out, &["left", "right"]);
}

#[test]
fn test_sealed_branches_can_be_mapped_without_else() {
    let out = run_prints(
        r#"
        sealed class Packet {
            class Text : Packet()
            class Number : Packet()
        }

        fun describe(packet: Packet): String {
            return when (packet) {
                is Packet.Text -> "text"
                is Packet.Number -> "number"
            }
        }

        fun main() {
            println(describe(Packet.Text()))
            println(describe(Packet.Number()))
        }
    "#,
    );
    assert_eq!(out, &["text", "number"]);
}

#[test]
fn test_state_shape_preserved_in_when_mapping() {
    let out = run_prints(
        r#"
        sealed class Variant {
            class A(val id: Int) : Variant()
            class B(val name: String) : Variant()
        }

        fun map(variant: Variant): String {
            return when (variant) {
                is Variant.A -> variant.id.toString()
                is Variant.B -> variant.name
            }
        }

        fun main() {
            println(map(Variant.A(4)))
            println(map(Variant.B("ok")))
        }
    "#,
    );
    assert_eq!(out, &["4", "ok"]);
}

#[test]
fn test_sealed_when_with_object_and_data_leaf_variants() {
    let out = run_prints(
        r#"
        sealed class Token {
            data class Named(val label: String) : Token()
            class Number(val value: Int) : Token()
            object Idle : Token()
        }

        fun classify(token: Token): String {
            return when (token) {
                is Token.Named -> token.label
                is Token.Number -> token.value.toString()
                is Token.Idle -> "idle"
            }
        }

        fun main() {
            println(classify(Token.Named("ok")))
            println(classify(Token.Number(7)))
            println(classify(Token.Idle))
        }
    "#,
    );
    assert_eq!(out, &["ok", "7", "idle"]);
}

#[test]
fn test_sealed_when_over_nullable_token() {
    let out = run_prints(
        r#"
        sealed class MaybeToken {
            class Present(val value: String) : MaybeToken()
            object Missing : MaybeToken()
        }

        fun render(token: MaybeToken?): String {
            return when (token) {
                is MaybeToken.Present -> token.value
                is MaybeToken.Missing -> "none"
                null -> "null"
            }
        }

        fun main() {
            println(render(MaybeToken.Present("a")))
            println(render(MaybeToken.Missing))
            println(render(null))
        }
    "#,
    );
    assert_eq!(out, &["a", "none", "null"]);
}

#[test]
fn test_local_sealed_class_in_function_evaluates_all_branches() {
    let out = run_prints(
        r#"
        fun classify(flag: Boolean): String {
            sealed class LocalResult {
                class Yes(val label: String) : LocalResult()
                class No : LocalResult()
            }

            val result: LocalResult = if (flag) LocalResult.Yes("ok") else LocalResult.No()
            return when (result) {
                is LocalResult.Yes -> result.label
                is LocalResult.No -> "no"
            }
        }

        fun main() {
            println(classify(true))
            println(classify(false))
        }
    "#,
    );
    assert_eq!(out, &["ok", "no"]);
}

#[test]
fn test_sealed_generic_shape_is_stable_across_branches() {
    let out = run_prints(
        r#"
        sealed class Value {
            class Text(val value: String) : Value()
            class Count(val value: Int) : Value()
        }

        fun normalize(value: Value): String {
            return when (value) {
                is Value.Text -> value.value
                is Value.Count -> value.value.toString()
            }
        }

        fun main() {
            val list: List<Value> = listOf(Value.Text("x"), Value.Count(3))
            println(normalize(list[0]))
            println(normalize(list[1]))
        }
    "#,
    );
    assert_eq!(out, &["x", "3"]);
}

#[test]
fn test_nested_sealed_dispatch_keeps_payload_associations() {
    let out = run_prints(
        r#"
        sealed class Packet {
            class Left(val code: Int) : Packet()
            class Right(val label: String) : Packet()
        }

        sealed class Wrapper {
            class Item(val packet: Packet) : Wrapper()
            class Empty : Wrapper()
        }

        fun describe(wrapper: Wrapper): String {
            return when (wrapper) {
                is Wrapper.Empty -> "none"
                is Wrapper.Item -> when (wrapper.packet) {
                    is Packet.Left -> "L" + wrapper.packet.code
                    is Packet.Right -> "R" + wrapper.packet.label
                }
            }
        }

        fun main() {
            println(describe(Wrapper.Item(Packet.Left(2))))
            println(describe(Wrapper.Item(Packet.Right("ok"))))
            println(describe(Wrapper.Empty()))
        }
    "#,
    );
    assert_eq!(out, &["L2", "Rok", "none"]);
}

#[test]
fn test_sealed_subclasses_respect_object_singleton_instance() {
    let out = run_prints(
        r#"
        sealed class State {
            object Active : State()
            class Paused(val count: Int) : State()
        }

        fun render(state: State): String {
            return when (state) {
                is State.Active -> "active"
                is State.Paused -> "paused-" + state.count.toString()
            }
        }

        fun main() {
            println(render(State.Active))
            println(render(State.Paused(4)))
            println(render(State.Active))
        }
    "#,
    );
    assert_eq!(out, &["active", "paused-4", "active"]);
}

#[test]
fn test_sealed_types_in_sequences_keep_exhaustive_mapping() {
    let out = run_prints(
        r#"
        sealed class Action {
            class Append(val value: String) : Action()
            class Multiply(val value: Int) : Action()
        }

        fun emit(actions: List<Action>): String {
            return actions.joinToString(",") {
                when (it) {
                    is Action.Append -> it.value
                    is Action.Multiply -> "x" + it.value.toString()
                }
            }
        }

        fun main() {
            val actions = listOf(
                Action.Append("a"),
                Action.Multiply(2),
                Action.Append("b"),
            )
            println(emit(actions))
        }
    "#,
    );
    assert_eq!(out, &["a,x2,b"]);
}

#[test]
fn test_sealed_interface_dispatched_like_closed_world_protocol() {
    let out = run_prints(
        r#"
        sealed interface Transport

        class Bus : Transport
        class Train : Transport

        fun is_mass_transport(value: Transport): Boolean {
            return when (value) {
                is Bus -> true
                is Train -> true
            }
        }

        fun main() {
            println(is_mass_transport(Bus()))
            println(is_mass_transport(Train()))
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_when_on_sealed_interface_in_function_body_with_local_value() {
    let out = run_prints(
        r#"
        sealed interface Stage
        class Start : Stage
        class End : Stage

        fun stage_text(stage: Stage): String {
            val value: Stage = stage
            return when (value) {
                is Start -> "start"
                is End -> "end"
            }
        }

        fun main() {
            val start: Stage = Start()
            val end: Stage = End()
            println(stage_text(start))
            println(stage_text(end))
        }
    "#,
    );
    assert_eq!(out, &["start", "end"]);
}

#[test]
fn test_sealed_nested_structure_can_be_traversed_iteratively() {
    let out = run_prints(
        r#"
        sealed class Node {
            class Leaf(val value: Int) : Node()
            class Branch(val left: Node, val right: Node) : Node()
        }

        fun collect(node: Node): Int {
            return when (node) {
                is Node.Leaf -> 1
                is Node.Branch -> 1 + collect(node.left) + collect(node.right)
            }
        }

        fun main() {
            val root = Node.Branch(
                Node.Leaf(1),
                Node.Branch(Node.Leaf(2), Node.Leaf(3))
            )
            println(collect(root))
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}
