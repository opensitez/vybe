use vybe_runtime::{Chunk,Op,VM,Value};
use vybe_runtime::debugger::{DebugRequest,DebugResponse,DebugCommand,DebugEvent};
mod functions {
 use vybe_runtime::{Chunk,Op};
 pub fn create_function_chunk(name:&str,arity:u8)->Chunk {let mut c=Chunk::new(name);c.arity=arity;c}
 pub fn emit_ref_func(c:&mut Chunk,index:usize,captures:u8,line:u32){c.emit_op_u16(Op::REF_FUNC,u16::try_from(index).unwrap(),line);c.emit(captures,line);}
}
use std::sync::mpsc;
use std::time::Duration;
fn send(tx:&mpsc::Sender<DebugRequest>,command:DebugCommand)->mpsc::Receiver<DebugResponse>{let (reply,rx)=mpsc::channel();tx.send(DebugRequest{command,reply}).unwrap();rx}
fn body(value:i32)->Chunk {let mut c=functions::create_function_chunk("shared-scalar",1);c.direct_entry_has_captures=Some(false);c.emit_i32_const(value,7);c.emit_op(Op::RETURN,7);c}
fn run_reload(shared:bool)->(i32,i32){
 let (commands,command_rx)=mpsc::channel();let (events,event_rx)=mpsc::channel();
 let worker=std::thread::spawn(move||{
  let mut root=Chunk::new("root");functions::emit_ref_func(&mut root,2,0,7);root.emit_op(Op::RETURN,7);
  let target=body(7);let mut caller=body(7);
  if shared {caller.code.clear();caller.lines.clear();caller.emit_op_u16(Op::LOCAL_GET,0,7);caller.emit_direct_defined_call(1,&target,true,7).unwrap();
  }
  let chunks=vec![root,target,caller];
  let mut replacement=chunks.clone();replacement[1]=body(9);if !shared {replacement[2]=body(9);}
  let mut vm=VM::new();let function=vm.run(chunks).unwrap();vm.set_reload_hook(Box::new(move|_|Ok(replacement.clone())));
  vm.attach_debugger(command_rx,events,true);
  let first=vm.invoke(&function,&[Value::I32(42)]).unwrap();vm.detach_debugger();let next=vm.invoke(&function,&[Value::I32(42)]).unwrap();
  let (Value::I32(first),Value::I32(next))=(first,next) else {panic!("actual scalar results")};(first,next)
 });
 let skip=send(&commands,DebugCommand::SetSkipSystem{enabled:false});
 let event=event_rx.recv_timeout(Duration::from_secs(20)).unwrap();let DebugEvent::Paused{location,..}=event else {panic!("actual paused event: {event:?}")};assert_eq!(location.chunk_index,2);assert_eq!(location.ip,0);
 let skip_reply=skip.recv_timeout(Duration::from_secs(20)).unwrap();assert!(!matches!(skip_reply,DebugResponse::Error(_)),"skip-system reply: {skip_reply:?}");
 let result=send(&commands,DebugCommand::Reload).recv_timeout(Duration::from_secs(20)).unwrap();assert!(matches!(result,DebugResponse::Value(_)),"actual reload: {result:?}");
 let resume=send(&commands,DebugCommand::Detach);assert!(matches!(resume.recv_timeout(Duration::from_secs(20)).unwrap(),DebugResponse::Ok));
 worker.join().unwrap()
}
#[test]
fn active_original_inline_body_keeps_previous_value(){assert_eq!(run_reload(false),(7,9));}
#[test]
fn ordinary_active_direct_call_observes_reloaded_callee(){assert_eq!(run_reload(true),(9,9));}
