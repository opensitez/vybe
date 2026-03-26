pub mod decl;
pub mod expr;
pub mod stmt;
pub mod query;
pub mod xml;
pub mod core_types;

pub use decl::*;
pub use expr::*;
pub use stmt::*;
pub use query::*;
pub use xml::*;
pub use core_types::*;

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Program {
    pub declarations: Vec<Declaration>,
    pub statements: Vec<Statement>,
}
