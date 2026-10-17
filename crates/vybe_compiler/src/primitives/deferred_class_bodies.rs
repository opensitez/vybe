//! Opt-in owned compilation: defer method body ownership until every original
//! AST analysis has completed. This never contains an AST body clone.
use super::Compiler;
use vybe_ast::{ClassMember, ClassModifiers, Span, Statement, StmtKind};
use vybe_ast::class_normalize::{NormalClass, NormalMethod};
use std::collections::{HashMap, HashSet};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Destination { Instance(usize), Static(usize), Destructor }
#[derive(Clone, Copy, Debug)]
pub struct Donation { pub member: usize, pub destination: Destination }
pub struct DeferredClass { pub class: NormalClass, pub donations: Vec<Donation> }
pub type NormalizeFn = fn(Span, &str, &[String], &[String], &[ClassMember], &ClassModifiers) -> DeferredClass;

/// Validation-only admission accounting; no TLS access in normal builds.
#[cfg(feature = "service-verification")]
#[derive(Clone, Copy, Default, Debug)]
pub struct Statistics { pub methods: usize, pub statements: usize, pub statement_capacity_bytes: usize, pub installed_library_methods: usize, pub installed_library_statements: usize }
#[cfg(feature = "service-verification")]
thread_local! { static STATISTICS: std::cell::Cell<Statistics> = const { std::cell::Cell::new(Statistics { methods: 0, statements: 0, statement_capacity_bytes: 0, installed_library_methods: 0, installed_library_statements: 0 }) }; }
#[cfg(feature = "service-verification")]
pub fn take_statistics() -> Statistics { STATISTICS.with(|value| value.replace(Statistics::default())) }

// One coherent executable can compare the actual old normalized-body work
// while keeping production donation and structural admission identical. This
// oracle has no branch or TLS access outside the existing verification feature.
#[cfg(feature = "service-verification")]
thread_local! { static KEEP_INSTALLED_BODIES: std::cell::Cell<bool> = const {std::cell::Cell::new(false)}; }
#[cfg(feature = "service-verification")]
pub fn with_installed_body_baseline_for_validation<T>(run: impl FnOnce()->T) -> T {
    struct Restore(bool);
    impl Drop for Restore {fn drop(&mut self){KEEP_INSTALLED_BODIES.with(|value|value.set(self.0));}}
    let _restore=Restore(KEEP_INSTALLED_BODIES.with(|value|value.replace(true)));
    run()
}
#[cfg(feature = "service-verification")]
pub fn installed_metadata_admission_for_validation(body:&[Statement],case_sensitive:bool)->bool {
    has_installed_library_metadata_headers(body,case_sensitive)
}

fn class_headers_allowed(modifiers: &ClassModifiers, members: &[ClassMember]) -> bool {
    !modifiers.is_partial && matches!(modifiers.kind, vybe_ast::ClassKind::Class | vybe_ast::ClassKind::Interface)
        && !members.iter().any(|m| matches!(m, ClassMember::Augment(_) | ClassMember::NestedType(_)))
}
fn candidate_headers(modifiers: &ClassModifiers, members: &[ClassMember]) -> bool {
    modifiers.kind == vybe_ast::ClassKind::Class
        && !matches!(modifiers.linkage, vybe_ast::ClassLinkage::LibraryReference { .. })
        && members.iter().any(|member| matches!(member, ClassMember::Method(statement)
            if matches!(&statement.kind, StmtKind::FunctionDecl { body, .. } if !body.is_empty())))
}
/// Header-only admission, shared with the consuming owner. In particular,
/// installed library references do not make functions-only includes pay for a
/// second deep AST scan. This does not supply the independent-lowering proof.
pub(crate) fn has_candidate_class_headers(body: &[Statement]) -> bool {
    fn visit(body: &[Statement], found: &mut bool) -> bool {
        for statement in body {
            match &statement.kind {
                StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => if !visit(body, found) { return false; },
                StmtKind::ClassDecl { members, modifiers, .. } => {
                    if !class_headers_allowed(modifiers, members) { return false; }
                    *found |= candidate_headers(modifiers, members);
                }
                StmtKind::StructDecl { .. } | StmtKind::ModuleDecl { .. } => return false,
                _ => {},
            }
        }
        true
    }
    let mut found=false;
    visit(body, &mut found) && found
}

/// Borrow-free metadata admission for installed-library-only units. Inspect
/// statement headers first; installed method bodies are not traversed. Then
/// inspect source expression scopes, rejecting anonymous classes and nested
/// declarations in closures. No source Vec is donated in this mode, so the
/// separate independent-lowering ownership proof is unnecessary.
pub(crate) fn has_installed_library_metadata_headers(body: &[Statement], case_sensitive: bool) -> bool {
    installed_library_metadata_safe(body,case_sensitive,false,true)
}
/// Select ownership without authorizing omission of any body. Source classes
/// remain excluded here, so a failed Owner proof cannot donate their bodies.
/// Only Owner performs the expression-scope proof before actual elision.
pub(crate) fn has_installed_library_owned_candidate(body: &[Statement], case_sensitive: bool) -> bool {
    installed_library_metadata_safe(body,case_sensitive,false,false)
}
fn installed_library_metadata_safe(body:&[Statement],case_sensitive:bool,allow_source_classes:bool,check_expression_scopes:bool)->bool {
    struct Headers { installed: bool, seen: HashSet<String>, case_sensitive: bool, allow_source_classes: bool }
    impl Headers {
        fn qualified(&self, namespace: &str, name: &str) -> String {
            let name=name.replace('\\',".");
            let key=if namespace.is_empty() || name.contains('.') {name} else {format!("{namespace}.{name}")};
            if self.case_sensitive {key} else {key.to_lowercase()}
        }
        fn body(&mut self, body: &[Statement], namespace: &str, nested: bool) -> bool {
            body.iter().all(|statement|self.statement(statement,namespace,nested))
        }
        fn statement(&mut self, statement: &Statement, namespace: &str, nested: bool) -> bool {
            match &statement.kind {
                StmtKind::NamespaceDecl {name,body} => {
                    let next=self.qualified(namespace,name);
                    self.body(body,&next,nested)
                },
                StmtKind::Block(body)=>self.body(body,namespace,nested),
                StmtKind::ClassDecl {name,members,modifiers,..} => {
                    if nested || !class_headers_allowed(modifiers,members)
                        || !self.seen.insert(self.qualified(namespace,name)) {return false;}
                    match (&modifiers.kind,&modifiers.linkage) {
                        (vybe_ast::ClassKind::Class,vybe_ast::ClassLinkage::LibraryReference{..})=>{
                            self.installed=true;true
                        },
                        (vybe_ast::ClassKind::Interface,vybe_ast::ClassLinkage::LibraryReference{..})=>true,
                        (vybe_ast::ClassKind::Interface,_)=>{
                            members.iter().all(|member|match member {
                                ClassMember::Method(statement)=>self.statement(statement,namespace,true),
                                ClassMember::Constructor{body,..}=>self.body(body,namespace,true),
                                ClassMember::Property{getter,setter,..}=>getter.as_ref().is_none_or(|body|self.body(body,namespace,true))
                                    && setter.as_ref().is_none_or(|setter|self.body(&setter.body,namespace,true)),
                                ClassMember::Field{..}=>true,
                                _=>false,
                            })
                        },
                        (vybe_ast::ClassKind::Class,_) if self.allow_source_classes=>{
                            members.iter().all(|member|match member {
                                ClassMember::Method(statement)=>self.statement(statement,namespace,true),
                                ClassMember::Constructor{body,..}=>self.body(body,namespace,true),
                                ClassMember::Property{getter,setter,..}=>getter.as_ref().is_none_or(|body|self.body(body,namespace,true))
                                    && setter.as_ref().is_none_or(|setter|self.body(&setter.body,namespace,true)),
                                ClassMember::Field{..}=>true,
                                _=>false,
                            })
                        },
                        _=>false,
                    }
                },
                StmtKind::FunctionDecl {body,..}|StmtKind::DoWhile {body,..}=>self.body(body,namespace,true),
                StmtKind::If {then_body,elifs,else_body,..}=>self.body(then_body,namespace,true)
                    && elifs.iter().all(|(_,body)|self.body(body,namespace,true))
                    && else_body.as_ref().is_none_or(|body|self.body(body,namespace,true)),
                StmtKind::For {init,body,..}=>init.as_ref().is_none_or(|s|self.statement(s,namespace,true))
                    && self.body(body,namespace,true),
                StmtKind::ForIn {body,else_body,..}|StmtKind::While {body,else_body,..}=>self.body(body,namespace,true)
                    && else_body.as_ref().is_none_or(|body|self.body(body,namespace,true)),
                StmtKind::Switch {cases,default,..}=>cases.iter().all(|case|self.body(&case.body,namespace,true))
                    && default.as_ref().is_none_or(|body|self.body(body,namespace,true)),
                StmtKind::Try {body,catches,else_body,finally}=>self.body(body,namespace,true)
                    && catches.iter().all(|catch|self.body(&catch.body,namespace,true))
                    && else_body.as_ref().is_none_or(|body|self.body(body,namespace,true))
                    && finally.as_ref().is_none_or(|body|self.body(body,namespace,true)),
                StmtKind::Labeled {body,..}=>self.statement(body,namespace,true),
                StmtKind::Empty|StmtKind::ScopeDecl{..}|StmtKind::Echo(_)|StmtKind::Exit{..}|StmtKind::Delete(_)
                |StmtKind::Assert{..}|StmtKind::GoTo(_)|StmtKind::Label(_)|StmtKind::Expr(_)|StmtKind::Directive{..}
                |StmtKind::VarDecl{..}|StmtKind::InterfaceDecl{..}|StmtKind::Return(_)|StmtKind::Break(_)
                |StmtKind::Continue(_)|StmtKind::Throw{..}|StmtKind::Assign{..}|StmtKind::CompoundAssign{..}=>true,
                _=>false,
            }
        }
    }
    let mut headers=Headers {installed:false,seen:HashSet::new(),case_sensitive,allow_source_classes};
    if !headers.body(body,"",false) || !headers.installed {return false;}
    if !check_expression_scopes { return true; }
    fn source_scopes_safe(body:&[Statement],headers:&mut Headers)->bool {
        for statement in body {
            match &statement.kind {
                // Those bodies will not be emitted; no scope inside them can
                // replace a normalized source definition during this compile.
                StmtKind::ClassDecl{modifiers,..} if matches!(modifiers.linkage,vybe_ast::ClassLinkage::LibraryReference{..})=>{},
                StmtKind::NamespaceDecl{body,..}|StmtKind::Block(body)=>{
                    if !source_scopes_safe(body,headers){return false;}
                },
                _=>{
                    let mut safe=true;
                    statement.walk_exprs(&mut |expression|match &expression.kind {
                        vybe_ast::ExprKind::ClassExpr{..}=>safe=false,
                        vybe_ast::ExprKind::FunctionExpr(declaration)=>{
                            safe &= headers.statement(declaration,"",true);
                        },
                        vybe_ast::ExprKind::Lambda{body:vybe_ast::LambdaBody::Block(body),..}=>{
                            safe &= headers.body(body,"",true);
                        },
                        _=>{},
                    });
                    if !safe{return false;}
                },
            }
        }
        true
    }
    source_scopes_safe(body,&mut headers)
}

struct Pending { key: String, path: Vec<usize>, donations: Vec<Donation> }
pub(super) struct Owner {
    normalize: NormalizeFn,
    paths: HashMap<String, Vec<usize>>,
    pending: Vec<Pending>,
}

fn method(class: &NormalClass, target: Destination) -> Option<&NormalMethod> {
    match target {
        Destination::Instance(i) => class.instance_methods.get(i),
        Destination::Static(i) => class.static_methods.get(i),
        Destination::Destructor => class.destructor.as_ref(),
    }
}
fn method_mut(class: &mut NormalClass, target: Destination) -> Option<&mut NormalMethod> {
    match target {
        Destination::Instance(i) => class.instance_methods.get_mut(i),
        Destination::Static(i) => class.static_methods.get_mut(i),
        Destination::Destructor => class.destructor.as_mut(),
    }
}
fn source_body(members: &[ClassMember], index: usize) -> Option<&Vec<Statement>> {
    let ClassMember::Method(statement) = members.get(index)? else { return None; };
    let StmtKind::FunctionDecl { body, .. } = &statement.kind else { return None; };
    Some(body)
}
fn source_body_mut(members: &mut [ClassMember], index: usize) -> Option<&mut Vec<Statement>> {
    let ClassMember::Method(statement) = members.get_mut(index)? else { return None; };
    let StmtKind::FunctionDecl { body, .. } = &mut statement.kind else { return None; };
    Some(body)
}
fn members<'a>(body: &'a [Statement], path: &[usize]) -> Option<&'a [ClassMember]> {
    let (first, rest) = path.split_first()?;
    let statement = body.get(*first)?;
    if rest.is_empty() {
        let StmtKind::ClassDecl { members, .. } = &statement.kind else { return None; };
        return Some(members);
    }
    match &statement.kind {
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => members(body, rest),
        _ => None,
    }
}
fn members_mut<'a>(body: &'a mut [Statement], path: &[usize]) -> Option<&'a mut [ClassMember]> {
    let (first, rest) = path.split_first()?;
    let statement = body.get_mut(*first)?;
    if rest.is_empty() {
        let StmtKind::ClassDecl { members, .. } = &mut statement.kind else { return None; };
        return Some(members);
    }
    match &mut statement.kind {
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => members_mut(body, rest),
        _ => None,
    }
}
impl Owner {
    /// Headers only. Every nested body remains untouched for the existing
    /// address/reference/atomic/reflection/rest passes. The caller separately
    /// supplies the existing independent-lowering admission receipt.
    pub(super) fn new(compiler: &Compiler, normalize: NormalizeFn, body: &[Statement]) -> Option<Self> {
        fn visit(compiler: &Compiler, body: &[Statement], namespace: Option<&str>, path: &mut Vec<usize>,
            paths: &mut HashMap<String, Vec<usize>>, seen: &mut HashSet<String>, elide_installed:bool) -> Option<()> {
            for (index, statement) in body.iter().enumerate() {
                path.push(index);
                match &statement.kind {
                    StmtKind::NamespaceDecl { name, body } => {
                        let member = compiler.canon(name).replace('\\', ".");
                        let qualified = match namespace {
                            Some(prefix) if !prefix.is_empty() && !member.is_empty() => format!("{prefix}.{member}"),
                            Some(prefix) if !prefix.is_empty() => prefix.to_string(),
                            _ => member,
                        };
                        visit(compiler, body, Some(&qualified), path, paths, seen,elide_installed)?;
                    }
                    StmtKind::Block(body) => visit(compiler, body, namespace, path, paths, seen,elide_installed)?,
                    StmtKind::ClassDecl { name, members, modifiers, .. } => {
                        // No augmentation can copy a not-yet-filled body. No
                        // partial/duplicate key may replace a pending target.
                        if !class_headers_allowed(modifiers, members) { return None; }
                        let member = compiler.canon(name).replace('\\', ".");
                        let key = match namespace {
                            Some(prefix) if !prefix.is_empty() && !member.contains('.') => format!("{prefix}.{member}"),
                            _ => member,
                        };
                        if !seen.insert(key.clone()) { return None; }
                        // Installed providers already own executable bodies. Within
                        // this proven augmentation-free owner, only their method
                        // metadata is required; source AST analyses remain intact.
                        if candidate_headers(modifiers, members)
                            || (elide_installed && modifiers.kind == vybe_ast::ClassKind::Class
                                && matches!(modifiers.linkage, vybe_ast::ClassLinkage::LibraryReference { .. })) {
                            paths.insert(key, path.clone());
                        }
                    }
                    StmtKind::StructDecl { .. } | StmtKind::ModuleDecl { .. } => return None,
                    _ => {},
                }
                path.pop();
            }
            Some(())
        }
        let mut paths = HashMap::new();
        let elide_installed=installed_library_metadata_safe(body,compiler.case_sensitive,true,true);
        visit(compiler, body, None, &mut Vec::new(), &mut paths, &mut HashSet::new(),elide_installed)?;
        (!paths.is_empty()).then_some(Self { normalize, paths, pending: Vec::new() })
    }
    pub(super) fn prepare(&mut self, span: Span, key: &str, parents: &[String], interfaces: &[String],
        members: &[ClassMember], modifiers: &ClassModifiers) -> Option<NormalClass> {
        let path = self.paths.get(key)?.clone();
        #[cfg(feature = "service-verification")]
        if matches!(modifiers.linkage,vybe_ast::ClassLinkage::LibraryReference{..})
            && KEEP_INSTALLED_BODIES.with(|value|value.get()) {return None;}
        let deferred = (self.normalize)(span, key, parents, interfaces, members, modifiers);
        let mut sources = HashSet::new();
        let mut targets = HashSet::new();
        // Validate every receipt before publishing metadata. A malformed hook
        // is declined while the original source bodies are still untouched.
        if !deferred.class.augmentations.is_empty() || deferred.donations.iter().any(|d|
            !sources.insert(d.member) || !targets.insert(d.destination)
            || source_body(members, d.member).is_none()
            || !method(&deferred.class, d.destination).is_some_and(|m| m.body.is_empty())) {
            return None;
        }
        if matches!(modifiers.linkage, vybe_ast::ClassLinkage::LibraryReference { .. }) {
            // Do not clone, donate, or clear the source body: reflection/rest
            // analyses still read it. Emission skips this validated installed
            // declaration, while constructor/property storage analyses retain
            // their ordinary full normalized views.
            #[cfg(feature = "service-verification")]
            STATISTICS.with(|statistics| {
                let mut value=statistics.get();
                value.installed_library_methods+=deferred.donations.len();
                value.installed_library_statements+=deferred.donations.iter()
                    .map(|d| source_body(members,d.member).unwrap().len()).sum::<usize>();
                statistics.set(value);
            });
        } else {
            self.pending.push(Pending { key: key.to_string(), path, donations: deferred.donations });
        }
        Some(deferred.class)
    }
    /// Borrowed compilation may publish installed-provider metadata, but must
    /// never take or clone original method bodies to satisfy a donation.
    pub(super) fn finish_metadata_only(self) -> Result<(), String> {
        if self.pending.is_empty() { Ok(()) }
        else { Err("borrowed class normalization attempted source-body donation".into()) }
    }
    pub(super) fn finish(self, compiler: &mut Compiler, body: &mut [Statement]) -> Result<(), String> {
        // Validate the entire transaction before taking the first source Vec.
        for pending in &self.pending {
            let source = members(body, &pending.path).ok_or("lost deferred class source")?;
            let class = compiler.normalized_classes.get(&pending.key).ok_or("lost deferred class target")?;
            for donation in &pending.donations {
                if source_body(source, donation.member).is_none()
                    || !method(class, donation.destination).is_some_and(|m| m.body.is_empty()) {
                    return Err("changed deferred class body receipt".into());
                }
            }
        }
        for pending in self.pending {
            let source = members_mut(body, &pending.path).unwrap();
            let class = compiler.normalized_classes.get_mut(&pending.key).unwrap();
            for donation in pending.donations {
                let body = source_body_mut(source, donation.member).unwrap();
                #[cfg(feature = "service-verification")]
                STATISTICS.with(|statistics| {
                    let mut value = statistics.get();
                    value.methods += 1; value.statements += body.len();
                    value.statement_capacity_bytes += body.capacity() * std::mem::size_of::<Statement>();
                    statistics.set(value);
                });
                method_mut(class, donation.destination).unwrap().body = std::mem::take(body);
            }
        }
        Ok(())
    }
}

impl Compiler {
    pub(super) fn normalize_class_declaration(&mut self, span: Span, key: &str,
        parents: &[String], interfaces: &[String], members: &[ClassMember],
        modifiers: &ClassModifiers, semantics: vybe_ast::ValueSemantics) -> Result<NormalClass, String> {
        let pending_before=self.deferred_class_bodies.as_ref().map_or(0,|owner|owner.pending.len());
        let supplied = self.deferred_class_bodies.as_mut().and_then(|owner|
            owner.prepare(span, key, parents, interfaces, members, modifiers));
        let result = super::class_normalize::emit::normalize_class_from_ast_with_override(
            self, span, key, parents, interfaces, members, modifiers, semantics, supplied);
        if result.is_err() {
            if let Some(owner)=self.deferred_class_bodies.as_mut() { owner.pending.truncate(pending_before); }
        }
        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn compiler() -> Compiler {
        Compiler::with_profile(crate::profile::parse_profile("[info]\nname='donation-test'\n[compiler]\n").unwrap())
    }
    fn declaration() -> Statement {
        Statement::new(StmtKind::ClassDecl {
            name: "Example".into(), parents: Vec::new(), interfaces: Vec::new(), decorators: Vec::new(),
            modifiers: ClassModifiers::default(),
            members: vec![ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                name: "run".into(), params: Vec::new(), return_type: None,
                body: vec![Statement::new(StmtKind::Empty), Statement::new(StmtKind::Empty)],
                modifiers: vybe_ast::Modifiers::default(), handles: Vec::new(),
                is_async: false, is_generator: false, is_sub: false,
            })))],
        })
    }
    fn header(span: Span, _: &str, _: &[String], _: &[String], _: &[ClassMember], _: &ClassModifiers) -> DeferredClass {
        let mut class = NormalClass::default();
        class.instance_methods.push(vybe_ast::class_normalize::build_normal_method(
            span, "run", "run", Vec::new(), None, Vec::new(),
            vybe_ast::class_normalize::Access::Public, false, false, false, vybe_ast::Modifiers::default()));
        DeferredClass { class, donations: vec![Donation { member: 0, destination: Destination::Instance(0) }] }
    }
    #[test]
    fn moves_original_allocation_after_original_analyses() {
        let mut compiler = compiler();
        let mut body = vec![declaration()];
        let original = source_body(members(&body, &[0]).unwrap(), 0).unwrap();
        let pointer = original.as_ptr(); let capacity = original.capacity();
        let mut owner = Owner::new(&compiler, header, &body).unwrap();
        let class = owner.prepare(Span::default(), &compiler.canon("Example"), &[], &[],
            members(&body, &[0]).unwrap(), &ClassModifiers::default()).unwrap();
        // The source body remains readable by original reflection/rest/scans.
        assert_eq!(source_body(members(&body, &[0]).unwrap(), 0).unwrap().as_ptr(), pointer);
        assert_eq!(source_body(members(&body, &[0]).unwrap(), 0).unwrap().len(), 2);
        assert!(class.instance_methods[0].body.is_empty());
        let key = compiler.canon("Example");
        compiler.normalized_classes.insert(key, class);
        owner.finish(&mut compiler, &mut body).unwrap();
        let donated = &compiler.normalized_classes[&compiler.canon("Example")].instance_methods[0].body;
        assert_eq!(donated.as_ptr(), pointer); assert_eq!(donated.capacity(), capacity); assert_eq!(donated.len(), 2);
        assert!(source_body(members(&body, &[0]).unwrap(), 0).unwrap().is_empty());
    }
    #[test]
    fn installed_reference_omits_normalized_method_without_taking_source() {
        let mut compiler=compiler();
        let mut installed=declaration();
        let StmtKind::ClassDecl {name,modifiers,..}=&mut installed.kind else {unreachable!()};
        *name="Installed".into();
        modifiers.linkage=vybe_ast::ClassLinkage::LibraryReference {identity:"test:installed:v1".into()};
        let modifiers=modifiers.clone();
        let mut body=vec![installed,declaration()];
        let original=source_body(members(&body,&[0]).unwrap(),0).unwrap();
        let pointer=original.as_ptr();let capacity=original.capacity();
        let mut owner=Owner::new(&compiler,header,&body).unwrap();
        let key=compiler.canon("Installed");
        let normalized=owner.prepare(Span::default(),&key,&[],&[],
            members(&body,&[0]).unwrap(),&modifiers).unwrap();
        assert!(normalized.instance_methods[0].body.is_empty());
        assert!(owner.pending.is_empty(),"installed body must never be donated");
        compiler.normalized_classes.insert(key.clone(),normalized);
        owner.finish(&mut compiler,&mut body).unwrap();
        let original=source_body(members(&body,&[0]).unwrap(),0).unwrap();
        assert_eq!(original.as_ptr(),pointer);assert_eq!(original.capacity(),capacity);assert_eq!(original.len(),2);
        assert!(compiler.normalized_classes[&key].instance_methods[0].body.is_empty());
    }
    #[test]
    fn borrowed_metadata_finishes_without_source_mutation() {
        let compiler=compiler();
        let mut installed=declaration();
        let StmtKind::ClassDecl {modifiers,..}=&mut installed.kind else {unreachable!()};
        modifiers.linkage=vybe_ast::ClassLinkage::LibraryReference {identity:"borrowed:provider".into()};
        let modifiers=modifiers.clone();
        let body=vec![installed];
        assert!(has_installed_library_metadata_headers(&body,compiler.case_sensitive));
        let snapshot=format!("{body:?}");
        let source=source_body(members(&body,&[0]).unwrap(),0).unwrap();
        let pointer=source.as_ptr();let capacity=source.capacity();
        let mut owner=Owner::new(&compiler,header,&body).unwrap();
        let class=owner.prepare(Span::default(),&compiler.canon("Example"),&[],&[],
            members(&body,&[0]).unwrap(),&modifiers).unwrap();
        assert!(class.instance_methods[0].body.is_empty());
        owner.finish_metadata_only().unwrap();
        let source=source_body(members(&body,&[0]).unwrap(),0).unwrap();
        assert_eq!(source.as_ptr(),pointer);assert_eq!(source.capacity(),capacity);
        assert_eq!(format!("{body:?}"),snapshot);
    }
    #[test]
    fn borrowed_metadata_rejects_unexpected_source_donation_without_mutation() {
        let compiler=compiler();let body=vec![declaration()];
        assert!(!has_installed_library_metadata_headers(&body,compiler.case_sensitive));
        let snapshot=format!("{body:?}");
        let mut owner=Owner::new(&compiler,header,&body).unwrap();
        owner.prepare(Span::default(),&compiler.canon("Example"),&[],&[],
            members(&body,&[0]).unwrap(),&ClassModifiers::default()).unwrap();
        assert!(owner.finish_metadata_only().is_err());
        assert_eq!(format!("{body:?}"),snapshot);
    }
    #[test]
    fn installed_only_admission_rejects_nested_and_duplicate_definitions() {
        let mut installed=declaration();
        let StmtKind::ClassDecl {name,modifiers,..}=&mut installed.kind else {unreachable!()};
        *name="Installed".into();
        modifiers.linkage=vybe_ast::ClassLinkage::LibraryReference {identity:"test:installed:v1".into()};
        assert!(has_installed_library_metadata_headers(std::slice::from_ref(&installed),false));
        assert!(!has_installed_library_metadata_headers(&[installed.clone(),installed.clone()],false));
        assert!(!has_installed_library_metadata_headers(&[installed.clone(),declaration()],false));
        let mut nested=match &declaration().kind {
            StmtKind::ClassDecl{members,..}=>match &members[0] {ClassMember::Method(s)=>(**s).clone(),_=>unreachable!()},
            _=>unreachable!(),
        };
        let StmtKind::FunctionDecl {body,..}=&mut nested.kind else {unreachable!()};
        *body=vec![installed.clone()];
        assert!(!has_installed_library_metadata_headers(&[installed.clone(),nested.clone()],false));
        let closure=Statement::new(StmtKind::Expr(vybe_ast::Expression::new(
            vybe_ast::ExprKind::FunctionExpr(Box::new(nested)))));
        assert!(!has_installed_library_metadata_headers(&[installed.clone(),closure.clone()],false));
        let hidden = [installed.clone(), closure.clone()];
        assert!(has_installed_library_owned_candidate(&hidden, false));
        assert!(Owner::new(&compiler(), header, &hidden).is_none(),
            "owned candidate is not permission to omit bodies or donate hidden declarations");
        let mut interface = declaration();
        let StmtKind::ClassDecl { modifiers, members, .. } = &mut interface.kind else { unreachable!() };
        modifiers.kind = vybe_ast::ClassKind::Interface;
        let ClassMember::Method(method) = &mut members[0] else { unreachable!() };
        let StmtKind::FunctionDecl { body, .. } = &mut method.kind else { unreachable!() };
        *body = vec![installed.clone()];
        assert!(!has_installed_library_metadata_headers(&[installed.clone(), interface], false),
            "source-interface bodies cannot hide replacement declarations");
        let mut empty_class = declaration();
        let StmtKind::ClassDecl { members, .. } = &mut empty_class.kind else { unreachable!() };
        members.clear();
        assert!(!has_installed_library_owned_candidate(&[installed.clone(), empty_class], false));
        // Even the independently admitted donating path must not elide a
        // library schema when an expression-contained class can replace it.
        let compiler=compiler();
        let owner=Owner::new(&compiler,header,&[installed,closure,declaration()]).unwrap();
        assert_eq!(owner.paths.len(),1,"unsafe expression scope preserves installed bodies");
    }
    #[test]
    fn missing_target_rejects_before_taking_any_source() {
        let compiler = compiler(); let mut body = vec![declaration()];
        let mut owner = Owner::new(&compiler, header, &body).unwrap();
        owner.prepare(Span::default(), &compiler.canon("Example"), &[], &[],
            members(&body, &[0]).unwrap(), &ClassModifiers::default()).unwrap();
        let mut compiler = compiler;
        assert!(owner.finish(&mut compiler, &mut body).is_err());
        assert_eq!(source_body(members(&body, &[0]).unwrap(), 0).unwrap().len(), 2);
    }
    #[test]
    fn partial_and_duplicate_keys_decline_before_normalization() {
        let compiler = compiler(); let first = declaration();
        assert!(Owner::new(&compiler, header, &[first.clone(), first.clone()]).is_none());
        let mut partial = first;
        let StmtKind::ClassDecl { modifiers, .. } = &mut partial.kind else { unreachable!() };
        modifiers.is_partial = true;
        assert!(Owner::new(&compiler, header, &[partial]).is_none());
    }
    #[test]
    fn header_probe_skips_functions_only_and_installed_libraries() {
        assert!(!has_candidate_class_headers(&[Statement::new(StmtKind::Empty)]));
        let ordinary=declaration();assert!(has_candidate_class_headers(std::slice::from_ref(&ordinary)));
        let mut installed=ordinary;
        let StmtKind::ClassDecl { modifiers, .. }=&mut installed.kind else { unreachable!() };
        modifiers.linkage=vybe_ast::ClassLinkage::LibraryReference {identity:"already-installed".into()};
        assert!(!has_candidate_class_headers(&[installed]));
    }

    #[test]
    fn interface_headers_remain_borrowed_without_blocking_ordinary_class() {
        let compiler=compiler();
        let mut interface=declaration();
        let StmtKind::ClassDecl {name,modifiers,..}=&mut interface.kind else {unreachable!()};
        *name="Contract".into();modifiers.kind=vybe_ast::ClassKind::Interface;
        assert!(!has_candidate_class_headers(std::slice::from_ref(&interface)));
        let ordinary=declaration();
        assert!(has_candidate_class_headers(&[interface.clone(),ordinary.clone()]));
        let owner=Owner::new(&compiler,header,&[interface.clone(),ordinary]).unwrap();
        assert_eq!(owner.paths.len(),1);
        assert!(owner.paths.contains_key(&compiler.canon("Example")));
        let StmtKind::ClassDecl {name,..}=&mut interface.kind else {unreachable!()};
        *name="Example".into();
        assert!(Owner::new(&compiler,header,&[interface,declaration()]).is_none(),"interface participates in duplicate guard");
    }

    #[test]
    fn malformed_callback_receipt_declines_with_original_body_intact() {
        fn invalid(span:Span,name:&str,p:&[String],i:&[String],m:&[ClassMember],mods:&ClassModifiers)->DeferredClass {
            let mut result=header(span,name,p,i,m,mods);result.donations[0].member=99;result
        }
        let compiler=compiler();let body=vec![declaration()];
        let mut owner=Owner::new(&compiler,invalid,&body).unwrap();
        assert!(owner.prepare(Span::default(),&compiler.canon("Example"),&[],&[],
            members(&body,&[0]).unwrap(),&ClassModifiers::default()).is_none());
        assert!(owner.pending.is_empty());
        assert_eq!(source_body(members(&body,&[0]).unwrap(),0).unwrap().len(),2);
    }

}
