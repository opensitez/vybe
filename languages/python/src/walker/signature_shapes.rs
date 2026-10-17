use super::*;
// Argument descriptor verdict only: no Expression values/default trees are materialized.
fn validate(params:&[Param],kinds:&[String],args:&[Argument])->Result<(),Expression>{
 debug_assert!(args.iter().all(|a|!a.spread));
 let has_var_kwargs=params.iter().enumerate().any(|(i,_)|python_signature_kind(params,kinds,i)=="VAR_KEYWORD");
 let has_var_positional=params.iter().enumerate().any(|(i,_)|python_signature_kind(params,kinds,i)=="VAR_POSITIONAL");
 let mut consumed=HashSet::<usize>::new();let positional_count=args.iter().filter(|a|a.name.is_none()).count();let mut positional_index=0;
 for(i,param)in params.iter().enumerate(){match python_signature_kind(params,kinds,i).as_str(){
  "VAR_POSITIONAL"=>{positional_index=positional_count;},"VAR_KEYWORD"=>{},
  "KEYWORD_ONLY"=>{if let Some((j,_))=args.iter().enumerate().find(|(j,a)|!consumed.contains(j)&&a.name.as_deref()==Some(param.name.as_str())){consumed.insert(j);}else if param.default.is_none(){return Err(py_raise_expr("TypeError",Some("missing required keyword-only argument")));}},
  "POSITIONAL_ONLY"=>{if args.iter().any(|a|a.name.as_deref()==Some(param.name.as_str())){return Err(py_raise_expr("TypeError",Some("positional-only argument passed as keyword argument")));}if positional_index<positional_count{positional_index+=1;}else if param.default.is_none(){return Err(py_raise_expr("TypeError",Some("missing required positional argument")));}},
  _=>{let keyword=args.iter().enumerate().find(|(j,a)|!consumed.contains(j)&&a.name.as_deref()==Some(param.name.as_str()));if positional_index<positional_count{positional_index+=1;if keyword.is_some(){return Err(py_raise_expr("TypeError",Some("multiple values for argument")));}}else if let Some((j,_))=keyword{consumed.insert(j);}else if param.default.is_none(){return Err(py_raise_expr("TypeError",Some("missing required positional argument")));}}
 }}
 if positional_index<positional_count&&!has_var_positional{return Err(py_raise_expr("TypeError",Some("too many positional arguments")));}
 if !has_var_kwargs&&args.iter().enumerate().any(|(j,a)|a.name.is_some()&&!consumed.contains(&j)){return Err(py_raise_expr("TypeError",Some("unexpected keyword argument")));}
 Ok(())
}
pub(super)fn rewrite(__w:&mut PyWalker,callee:&Expression,args:&[Argument],optional:bool)->Option<Expression>{
 let ExprKind::Ident(name)=&callee.kind else{return None;};
 // Resolve twice, in the original params-then-kinds order; no result is cached.
 let Some(params_name)=resolve_function_name(__w,name)else{return None;};
 if __w.py_defined_function_params.get(&params_name).is_none(){return None;}
 let kinds_name=resolve_function_name(__w,name);
 // Both resolvers are read-only. Reborrow metadata after the second resolver to avoid
 // holding an immutable walker borrow across its original &mut signature.
 let params=__w.py_defined_function_params.get(&params_name).expect("read-only signature resolution");
 let kinds=kinds_name.as_ref().and_then(|key|__w.py_defined_function_param_kinds.get(key)).map(Vec::as_slice).unwrap_or(&[]);
 if params.is_empty(){return None;}
 if args.iter().all(|a|!a.spread){
  let had_named=args.iter().any(|a|a.name.is_some());let positional_count=args.iter().filter(|a|a.name.is_none()).count();
  let has_var_positional=params.iter().enumerate().any(|(i,_)|python_signature_kind(&params,&kinds,i)=="VAR_POSITIONAL");
  let has_var_keyword=params.iter().enumerate().any(|(i,_)|python_signature_kind(&params,&kinds,i)=="VAR_KEYWORD");
  let nonterminal=params.iter().enumerate().any(|(i,_)|python_signature_kind(&params,&kinds,i)=="VAR_POSITIONAL"&&kinds.iter().skip(i+1).any(|k|k!="VAR_KEYWORD"));
  if !had_named&&(has_var_positional||has_var_keyword)&&positional_count==params.len(){return None;}
  if !nonterminal&&!has_var_keyword{return match validate(&params,&kinds,args){Ok(())=>{None},Err(e)=>{Some(e)}};}
 }
 let params=params.to_vec();let kinds=kinds.to_vec();let args=args.to_vec();
    let had_spread = args.iter().any(|arg| arg.spread);
    let had_named_or_spread = args.iter().any(|arg| arg.name.is_some() || arg.spread);
    let expanded_args = python_expand_literal_spreads(args.clone()).unwrap_or(args.clone());
    let positional_count = expanded_args
        .iter()
        .filter(|arg| arg.name.is_none() && !arg.spread)
        .count();
    let has_var_positional = params
        .iter()
        .enumerate()
        .any(|(idx, _)| python_signature_kind(&params, &kinds, idx) == "VAR_POSITIONAL");
    let has_var_keyword = params
        .iter()
        .enumerate()
        .any(|(idx, _)| python_signature_kind(&params, &kinds, idx) == "VAR_KEYWORD");
    let has_nonterminal_var_positional = params.iter().enumerate().any(|(idx, _)| {
        python_signature_kind(&params, &kinds, idx) == "VAR_POSITIONAL"
            && kinds.iter().skip(idx + 1).any(|kind| kind != "VAR_KEYWORD")
    });
    if !had_named_or_spread
        && (has_var_positional || has_var_keyword)
        && positional_count == params.len()
    {
        return None;
    }
    if !has_nonterminal_var_positional && !has_var_keyword {
        return match python_signature_call_args(&params, &kinds, expanded_args.clone()) {
            Err(err) => Some(err),
            Ok(_) if had_spread => {
                let args = expanded_args
                    .into_iter()
                    .map(|mut arg| {
                        arg.value = desugar_member_reads(__w, arg.value);
                        arg
                    })
                    .collect();
                Some(Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args,
                    optional,
                }))
            }
            Ok(_) => None,
        };
    }
    match python_signature_call_args(&params, &kinds, args) {
        Ok(args) => {
            let args = args
                .into_iter()
                .map(|mut arg| {
                    arg.value = desugar_member_reads(__w, arg.value);
                    arg
                })
                .collect();
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(callee.clone()),
                args,
                optional,
            }))
        }
        Err(err) => Some(err),
    }

}

#[cfg(test)]
#[path = "signature_shapes_tests.rs"]
mod tests;
