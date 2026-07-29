//! Codepoint classification and mapping helpers.
//!
//! These functions operate on integer character/codepoint expressions and build
//! common AST. C `ctype.h` maps here for ASCII-compatible semantics; other
//! languages can reuse the same primitive where their rules match.

use vybe_ast::{BinOp, ExprKind, Expression, Literal, UnaryOp};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

/// `c >= lo && c <= hi` — inclusive integer range check.
pub fn int_range(c: Expression, lo: i64, hi: i64) -> Expression {
    let ge = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(lit(lo)),
    });
    let le = e(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(c),
        right: Box::new(lit(hi)),
    });
    e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(ge),
        right: Box::new(le),
    })
}

/// Normalize a boolean expression to C-like int semantics: `expr ? 1 : 0`.
pub fn bool_to_int(b: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(b),
        then: Box::new(lit(1)),
        else_: Box::new(lit(0)),
    })
}

pub fn c_isalpha(c: Expression) -> Expression {
    let upper = int_range(c.clone(), 65, 90);
    let lower = int_range(c, 97, 122);
    bool_to_int(e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(upper),
        right: Box::new(lower),
    }))
}

pub fn c_isdigit(c: Expression) -> Expression {
    bool_to_int(int_range(c, 48, 57))
}

pub fn c_isalnum(c: Expression) -> Expression {
    let alpha = c_isalpha(c.clone());
    let digit = c_isdigit(c);
    e(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(alpha),
        right: Box::new(digit),
    })
}

pub fn c_isspace(c: Expression) -> Expression {
    let sp = e(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c.clone()),
        right: Box::new(lit(32)),
    });
    let ctrl = int_range(c, 9, 13);
    bool_to_int(e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(sp),
        right: Box::new(ctrl),
    }))
}

pub fn c_isupper(c: Expression) -> Expression {
    bool_to_int(int_range(c, 65, 90))
}

pub fn c_islower(c: Expression) -> Expression {
    bool_to_int(int_range(c, 97, 122))
}

pub fn c_isxdigit(c: Expression) -> Expression {
    let dig = int_range(c.clone(), 48, 57);
    let uf = int_range(c.clone(), 65, 70);
    let lf = int_range(c, 97, 102);
    bool_to_int(e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(dig),
        right: Box::new(e(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(uf),
            right: Box::new(lf),
        })),
    }))
}

pub fn c_iscntrl(c: Expression) -> Expression {
    let lt32 = e(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(c.clone()),
        right: Box::new(lit(32)),
    });
    let eq127 = e(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c),
        right: Box::new(lit(127)),
    });
    e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(lt32),
        right: Box::new(eq127),
    })
}

pub fn c_isprint(c: Expression) -> Expression {
    let ge32 = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(lit(32)),
    });
    let lt127 = e(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(c),
        right: Box::new(lit(127)),
    });
    e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(ge32),
        right: Box::new(lt127),
    })
}

pub fn c_ispunct(c: Expression) -> Expression {
    let ge33 = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(lit(33)),
    });
    let le126 = e(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(c.clone()),
        right: Box::new(lit(126)),
    });
    let not_an = e(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(c_isalnum(c)),
    });
    e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(e(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(ge33),
            right: Box::new(le126),
        })),
        right: Box::new(not_an),
    })
}

pub fn c_isgraph(c: Expression) -> Expression {
    let ge33 = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(lit(33)),
    });
    let le126 = e(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(c),
        right: Box::new(lit(126)),
    });
    e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(ge33),
        right: Box::new(le126),
    })
}

pub fn c_isblank(c: Expression) -> Expression {
    let sp = e(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c.clone()),
        right: Box::new(lit(32)),
    });
    let tab = e(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c),
        right: Box::new(lit(9)),
    });
    e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(sp),
        right: Box::new(tab),
    })
}

pub fn c_toupper(c: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(int_range(c.clone(), 97, 122)),
        then: Box::new(e(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(c.clone()),
            right: Box::new(lit(32)),
        })),
        else_: Box::new(c),
    })
}

pub fn c_tolower(c: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(int_range(c.clone(), 65, 90)),
        then: Box::new(e(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(c.clone()),
            right: Box::new(lit(32)),
        })),
        else_: Box::new(c),
    })
}
