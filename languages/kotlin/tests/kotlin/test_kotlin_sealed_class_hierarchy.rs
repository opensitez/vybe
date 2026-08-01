use crate::helpers::run_prints;

#[test]
fn test_sealed_hierarchy_is_exhaustively_matched() {
    let out = run_prints(
        r#"
        sealed class Node {
            data class Value(val n: Int) : Node()
            data class Negate(val child: Node) : Node()
            data class Sum(val left: Node, val right: Node) : Node()
        }

        fun eval(node: Node): Int = when (node) {
            is Node.Value -> node.n
            is Node.Negate -> -eval(node.child)
            is Node.Sum -> eval(node.left) + eval(node.right)
        }

        fun main() {
            val expr = Node.Sum(Node.Value(3), Node.Negate(Node.Value(2)))
            println(eval(expr))
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_sealed_types_prevent_external_extension() {
    let out = run_prints(
        r#"
        sealed class Response {
            object Ok : Response()
            data class Error(val code: Int) : Response()
        }

        fun message(response: Response): String = when (response) {
            is Response.Ok -> "ok"
            is Response.Error -> "err=" + response.code
        }

        fun main() {
            println(message(Response.Ok))
            println(message(Response.Error(5)))
        }
    "#,
    );
    assert_eq!(out, &["ok", "err=5"]);
}
