#[allow(dead_code, unused_comparisons)]
mod reference {
    include!("fixtures/native_tokens_reference.rs");
}
#[allow(dead_code, unused_comparisons)]
mod candidate {
    include!("fixtures/native_tokens_candidate.rs");
}
use grameo::{
    MatchOptions, Span,
    arena::{Arena, Id},
    builder::Builder,
    pratt::Fixity,
    program::Program,
};
#[derive(Debug)]
enum Node {
    Number(String),
    Unary(&'static str, Id),
    Binary(&'static str, Id, Id),
}
struct Ast {
    arena: Arena<Node>,
    expr: usize,
    number: usize,
    plus: usize,
    minus: usize,
    star: usize,
    power: usize,
    bang: usize,
    reject: bool,
    reductions: Vec<(usize, Fixity, Span)>,
}
impl Ast {
    fn new<P: Program>(p: &P, reject: bool) -> Self {
        Self {
            arena: Arena::default(),
            expr: p.rule_id("expr").unwrap(),
            number: p.rule_id("number").unwrap(),
            plus: p.rule_id("plus").unwrap(),
            minus: p.rule_id("minus").unwrap(),
            star: p.rule_id("star").unwrap(),
            power: p.rule_id("power").unwrap(),
            bang: p.rule_id("bang").unwrap(),
            reject,
            reductions: vec![],
        }
    }
    fn shape(&self, id: Id) -> String {
        match self.arena.get(id) {
            Node::Number(n) => n.clone(),
            Node::Unary(o, a) => format!("({o} {})", self.shape(*a)),
            Node::Binary(o, a, b) => format!("({o} {} {})", self.shape(*a), self.shape(*b)),
        }
    }
    fn shapes(&self) -> Vec<String> {
        self.arena
            .values()
            .iter()
            .map(|&id| self.shape(id))
            .collect()
    }
}
impl Builder for Ast {
    fn supports_choice_heads(&self) -> bool {
        true
    }
    fn observes_rule(&self, r: usize) -> bool {
        r == self.number
    }
    fn opaque_rule(&self, _: usize) -> bool {
        true
    }
    fn supports_pratt(&self, r: usize) -> bool {
        r == self.expr
    }
    fn accepts_pratt(&self, _: usize, _: Span) -> bool {
        !self.reject
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, m: usize) {
        self.arena.rollback(m)
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, r: usize, s: Span, source: &str) {
        if r == self.number {
            let n = self
                .arena
                .alloc(Node::Number(source[s.start..s.end].to_owned()));
            self.arena.push(n);
        }
    }
    fn reduce(&mut self, r: usize, f: Fixity, s: Span) {
        self.reductions.push((r, f, s));
        // This test binding's reject mode deliberately requests PEG recapture
        // and does not construct operator nodes from rejected nested islands.
        if self.reject {
            return;
        }
        let op = if r == self.plus {
            "+"
        } else if r == self.minus && f == Fixity::Prefix {
            "neg"
        } else if r == self.minus {
            "-"
        } else if r == self.star {
            "*"
        } else if r == self.power {
            "^"
        } else if r == self.bang {
            "!"
        } else {
            panic!("unknown operator")
        };
        let right = self.arena.pop().unwrap();
        let n = match f {
            Fixity::Prefix | Fixity::Postfix => Node::Unary(op, right),
            _ => Node::Binary(op, self.arena.pop().unwrap(), right),
        };
        let id = self.arena.alloc(n);
        self.arena.push(id);
    }
}
fn snapshot(t: grameo::ParseTree<'_, '_>) -> (usize, Vec<(usize, Span, usize)>) {
    let mut out = vec![];
    let mut pending: Vec<_> = t.pairs().map(|p| (p, 0)).collect();
    while let Some((p, d)) = pending.pop() {
        out.push((
            match p.as_rule() {
                grameo::CaptureRule::User(r) => r,
                _ => usize::MAX,
            },
            p.as_span(),
            d,
        ));
        pending.extend(p.into_inner().map(|p| (p, d + 1)));
    }
    (t.consumed(), out)
}
#[test]
fn native_tokens_and_choice_dispatch_preserve_generic_pratt_ast_and_rollback() {
    let mut cases = 0;
    for entry in ["program", "probe", "fallback"] {
        for source in [
            "1+2*3",
            "(1+2)*3",
            "1^2^3",
            "-1+2",
            "1!!",
            "+",
            "1+",
            "(1",
            "1+2 x",
            "1+2",
            "1é",
            "100_000",
            "12 + 34 #bad",
        ] {
            for reject in [false, true] {
                for work in [0, 8, 32, 128, 512, usize::MAX] {
                    for depth in [0, 1, 2, 3, 8, 32] {
                        let mut a = Ast::new(&reference::Parser, reject);
                        let mut b = Ast::new(&candidate::Parser, reject);
                        let options = MatchOptions {
                            max_steps: work,
                            max_rule_depth: depth,
                        };
                        let x = grameo::engine::build_compact_capture_program(
                            &reference::Parser,
                            entry,
                            source,
                            options,
                            &mut a,
                        )
                        .map(snapshot);
                        let y = grameo::engine::build_compact_capture_program(
                            &candidate::Parser,
                            entry,
                            source,
                            options,
                            &mut b,
                        )
                        .map(snapshot);
                        assert_eq!(
                            x, y,
                            "entry={entry} source={source:?} reject={reject} work={work} depth={depth}"
                        );
                        assert_eq!(a.shapes(), b.shapes());
                        assert_eq!(a.reductions, b.reductions);
                        if x.is_err() {
                            assert!(a.arena.values().is_empty());
                            assert!(b.arena.values().is_empty());
                        }
                        cases += 1;
                    }
                }
            }
        }
    }
    println!("exact generic Pratt AST/capture/reduction/error/rollback cases={cases}");
    let mut a = Ast::new(&candidate::Parser, false);
    grameo::engine::build_compact_capture_program(
        &candidate::Parser,
        "program",
        "1+2*3",
        MatchOptions::default(),
        &mut a,
    )
    .unwrap();
    assert_eq!(a.shapes(), ["(+ 1 (* 2 3))"]);
}
