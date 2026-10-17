//! Owned member chains; contextual rewriting follows the original postfix order.
use super::*;

pub(super) fn eager_candidate(start: usize, source: &str) -> bool {
    let bytes=source.as_bytes();
    let mut cursor=start;
    if !bytes.get(cursor).is_some_and(|b|b.is_ascii_alphabetic() || *b==b'_') {return false;}
    cursor+=1;
    while bytes.get(cursor).is_some_and(|b|b.is_ascii_alphanumeric() || *b==b'_') {cursor+=1;}
    let mut member=false;
    loop {
        while bytes.get(cursor).is_some_and(|b|matches!(b,b' '|b'\t'|b'\r')) {cursor+=1;}
        if bytes.get(cursor)!=Some(&b'.') {
            return member && !matches!(bytes.get(cursor),Some(b'('|b'['));
        }
        cursor+=1;
        while bytes.get(cursor).is_some_and(|b|matches!(b,b' '|b'\t'|b'\r')) {cursor+=1;}
        if !bytes.get(cursor).is_some_and(|b|b.is_ascii_alphabetic() || *b==b'_') {return false;}
        cursor+=1;
        while bytes.get(cursor).is_some_and(|b|b.is_ascii_alphanumeric() || *b==b'_') {cursor+=1;}
        member=true;
    }
}

pub(super) fn reduce(ast: &mut Ast<'_>, op: grameo::Span) {
    let Some(source)=ast.source else {return;};
    let Some(id)=ast.arena.values().last().copied() else {return;};
    let Value::Ready{metadata,..}=ast.arena.get(id) else {return;};
    if metadata.raw.end>op.start {return;}
    let Some(field)=source[op.start..op.end].trim().strip_prefix('.').map(str::trim) else {return;};
    ast.arena.edit(id,|value| {
        let Value::Ready{mut expression,mut metadata,mut replay}=std::mem::replace(value,Value::Depth(0)) else {unreachable!()};
        let previous=Undo::Member(metadata,expression.span);
        if metadata.layer==7 {
            expression.span=Span::default();
            *replay.last_mut().unwrap().span()=expression.span;
        }
        let raw=grameo::Span{start:metadata.raw.start,end:op.end};
        let expression=Expression::with_span(ExprKind::Member{object:Box::new(expression),field:field.to_owned(),null_safe:false},span(raw));
        replay.push(Replay::Member{raw:op,span:expression.span});
        metadata.raw=raw;
        metadata.anchor=None;
        metadata.anchor_rule=Rule::postfix;
        metadata.contextual=true;
        metadata.sequence=false;
        metadata.layer=7;
        metadata.boundary=raw.end;
        *value=Value::Ready{expression,metadata,replay};
        previous
    },|value,undo| {
        let (Value::Ready{expression,mut replay,..},Undo::Member(metadata,previous_span))=(std::mem::replace(value,Value::Depth(0)),undo) else {unreachable!()};
        let ExprKind::Member{mut object,..}=expression.kind else {unreachable!()};
        assert!(matches!(replay.pop(),Some(Replay::Member{..})));
        object.span=previous_span;
        *replay.last_mut().unwrap().span()=previous_span;
        *value=Value::Ready{expression:*object,metadata,replay};
    });
}

pub(super) fn normalize(walker: &mut PyWalker, mut expression: Expression) -> Expression {
    let original_span=expression.span;
    if let ExprKind::Member{object,..}=&expression.kind
        && !matches!(object.kind,ExprKind::Member{..})
    {
        let ExprKind::Member{object,field,..}=expression.kind else {unreachable!()};
        let object=*object;
        let expression=lower_python_owned_member(walker,object,field);
        let mut expression=desugar_member_reads(walker,expression);
        expression.span=original_span;
        return expression;
    }
    let mut fields=Vec::new();
    while let ExprKind::Member{object,field,..}=expression.kind {
        fields.push(field);
        expression=*object;
    }
    let mut expression=expression;
    for field in fields.into_iter().rev() {
        expression=lower_python_owned_member(walker,expression,field);
    }
    let mut expression=desugar_member_reads(walker,expression);
    expression.span=original_span;
    expression
}

fn trivia(bytes: &[u8], cursor: &mut usize) {
    while bytes.get(*cursor).is_some_and(|b| matches!(b,b' '|b'\t'|b'\r')) { *cursor += 1; }
}
fn scalar_token(bytes: &[u8], cursor: &mut usize) -> bool {
    let Some(&first)=bytes.get(*cursor) else {return false;};
    if matches!(first,b'\''|b'"') {
        if bytes.get(*cursor+1)==Some(&first) && bytes.get(*cursor+2)==Some(&first) {return false;}
        *cursor+=1;
        while let Some(&b)=bytes.get(*cursor) {
            *cursor+=1;
            if b==first {return true;}
            if matches!(b,b'\n'|b'\r') {return false;}
            if b==b'\\' {if bytes.get(*cursor).is_none() {return false;} *cursor+=1;}
        }
        return false;
    }
    if first.is_ascii_alphabetic() || first==b'_' {
        *cursor+=1;
        while bytes.get(*cursor).is_some_and(|b|b.is_ascii_alphanumeric()||*b==b'_') {*cursor+=1;}
        return true;
    }
    if first.is_ascii_digit() {
        *cursor+=1;
        while let Some(&b)=bytes.get(*cursor) {
            if b.is_ascii_alphanumeric() || matches!(b,b'_'|b'.') {*cursor+=1;}
            else if matches!(b,b'+'|b'-') && matches!(bytes.get(*cursor-1),Some(b'e'|b'E')) {*cursor+=1;}
            else {break;}
        }
        return true;
    }
    false
}
// Admission is only a conservative hint. The opted-in grammar certifies tokens.
pub(super) fn candidate(start: usize, source: &str) -> bool {
    let bytes=source.as_bytes(); let mut cursor=start;
    if !bytes.get(cursor).is_some_and(|b|b.is_ascii_alphabetic()||*b==b'_') {return false;}
    if !scalar_token(bytes,&mut cursor) {return false;}
    let mut owned=false;
    loop {
        trivia(bytes,&mut cursor);
        match bytes.get(cursor) {
            Some(b'.') => {
                cursor+=1; trivia(bytes,&mut cursor);
                if !bytes.get(cursor).is_some_and(|b|b.is_ascii_alphabetic()||*b==b'_') {return false;}
                if !scalar_token(bytes,&mut cursor) {return false;}
                owned=true;
            }
            Some(b'(') => {
                cursor+=1; trivia(bytes,&mut cursor);
                if bytes.get(cursor)!=Some(&b')') {
                    loop {
                        if !scalar_token(bytes,&mut cursor) {return false;}
                        trivia(bytes,&mut cursor);
                        if bytes.get(cursor)!=Some(&b',') {break;}
                        cursor+=1; trivia(bytes,&mut cursor);
                        if bytes.get(cursor)==Some(&b')') {break;}
                    }
                }
                if bytes.get(cursor)!=Some(&b')') {return false;}
                cursor+=1; owned=true;
            }
            Some(b'[') => return false,
            _ => return owned,
        }
    }
}
fn scalar_rule(text:&str)->Rule {
    match text { "True"=>Rule::true_kw,"False"=>Rule::false_kw,"None"=>Rule::none_kw,
        _ if text.starts_with(['\'','"'])=>Rule::string_literal,
        _ if text.as_bytes().first().is_some_and(u8::is_ascii_digit)=>Rule::numeric_literal,
        _=>Rule::identifier }
}
pub(super) fn scalar_call_args(source:&str,raw:grameo::Span,base:usize)->Option<Vec<Argument>> {
    let bytes=source.as_bytes();let mut cursor=raw.start.checked_sub(base)?+1;
    let mut args=Vec::new();
    loop {
        trivia(bytes,&mut cursor); if bytes.get(cursor)==Some(&b')') {break;}
        let start=cursor;if !scalar_token(bytes,&mut cursor) {return None;}
        let kind=leaf(scalar_rule(&source[start..cursor]),&source[start..cursor])?;
        trivia(bytes,&mut cursor);
        args.push(Argument{name:None,value:Expression::with_span(kind,span(grameo::Span{start:start+base,end:cursor+base})),by_ref:false,spread:false});
        if bytes.get(cursor)!=Some(&b',') {break;}cursor+=1;
    }
    if bytes.get(cursor)!=Some(&b')') || cursor+base+1!=raw.end {return None;}
    Some(args)
}
pub(super) fn normalize_owned_calls(walker:&mut PyWalker,mut expression:Expression)->Expression {
    let mut probe=&expression;
    while let ExprKind::Member{object,..}=&probe.kind {probe=object;}
    if !matches!(probe.kind,ExprKind::Call{..}) {return normalize(walker,expression);}
    let original_span=expression.span;
    enum Step { Member(String),Call(Vec<Argument>) }
    let mut steps=Vec::new();
    loop {match expression.kind {
        ExprKind::Member{object,field,..}=>{steps.push(Step::Member(field));expression=*object;},
        ExprKind::Call{callee,args,..}=>{steps.push(Step::Call(args));expression=*callee;},
        _=>break,
    }}
    for step in steps.into_iter().rev() {match step {
        Step::Member(field)=>expression=lower_python_owned_member(walker,expression,field),
        Step::Call(mut args)=>{
            for arg in &mut args {let value=std::mem::replace(&mut arg.value,Expression::null());arg.value=normalize_contextual_expression(walker,value);}
            expression=if args.is_empty(){lower_python_owned_empty_call(walker,expression)}else{lower_python_owned_call_args(walker,expression,args)};
        }
    }}
    let mut expression=desugar_member_reads(walker,expression);expression.span=original_span;expression
}
