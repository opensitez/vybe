//! Merge complete declaration units before forward declaration registration.
use std::collections::HashMap;
use vybe_ast::{ClassMember, Statement, StmtKind};

struct ClassGroup {
    index: usize,
    has_constructor: bool,
}

pub(super) struct PartialClassPreparation {
    case_sensitive: bool,
    groups: HashMap<String, ClassGroup>,
    body: Vec<Statement>,
}

impl PartialClassPreparation {
    pub(super) fn new(case_sensitive: bool, capacity: usize) -> Self {
        Self { case_sensitive, groups: HashMap::new(), body: Vec::with_capacity(capacity) }
    }

    /// Match the legacy activated merge, including unflagged declarations of
    /// the same name. Only top-level classes participate; nested scopes do not.
    pub(super) fn consume(&mut self, statements: &[Statement]) {
        let key = |name: &str| if self.case_sensitive { name.to_owned() } else { name.to_lowercase() };
        for statement in statements {
            if let StmtKind::ClassDecl { name, members, parents, interfaces, .. } = &statement.kind {
                let name = key(name);
                if let Some(group) = self.groups.get_mut(&name) {
                    let StmtKind::ClassDecl { members: merged, parents: bases, interfaces: implemented, .. }
                        = &mut self.body[group.index].kind else { unreachable!() };
                    for member in members {
                        if matches!(member, ClassMember::Constructor { .. }) {
                            if group.has_constructor { continue; }
                            group.has_constructor = true;
                        }
                        merged.push(member.clone());
                    }
                    for parent in parents {
                        if !bases.iter().any(|existing| key(existing) == key(parent)) { bases.push(parent.clone()); }
                    }
                    for interface in interfaces {
                        if !implemented.iter().any(|existing| key(existing) == key(interface)) { implemented.push(interface.clone()); }
                    }
                    continue;
                }
                self.groups.insert(name, ClassGroup { index: self.body.len(),
                    has_constructor: members.iter().any(|member| matches!(member, ClassMember::Constructor { .. })) });
            }
            self.body.push(statement.clone());
        }
    }

    pub(super) fn finish(self) -> Vec<Statement> { self.body }
}

#[cfg(test)]
#[path = "declaration_preparation_tests.rs"]
mod tests;
