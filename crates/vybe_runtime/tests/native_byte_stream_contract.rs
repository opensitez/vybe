use std::cell::RefCell;
use std::rc::Rc;
use std::sync::Arc;
use vybe_runtime::component::ValType;
use vybe_runtime::event_loop::{EventLoop, FuturePhase};
use vybe_runtime::value::Value;

#[test]
fn byte_sink_requires_explicit_bytes_and_preserves_typed_elements() {
    let mut events = EventLoop::new();
    let bytes = events.create_stream();
    let future = events.create_future();
    for value in [
        Value::I32(255),
        Value::I64(256),
        Value::F64(1.5),
        Value::String(Arc::from("not bytes")),
    ] {
        assert!(events.stream_push(bytes, value).is_err());
    }
    events.stream_push_bytes(bytes, vec![0, 128, 255]);
    assert_eq!(events.stream_read_bytes(bytes, 1), [0]);
    let received = Rc::new(RefCell::new(Vec::new()));
    let eof = Rc::new(RefCell::new(0));
    let output = received.clone();
    let ended = eof.clone();
    events.set_byte_consumer(bytes, future, Box::new(move |chunk| {
        match chunk {
            Some(chunk) => output.borrow_mut().extend_from_slice(chunk),
            None => *ended.borrow_mut() += 1,
        }
        Ok(())
    })).unwrap();
    assert_eq!(events.future_states[&future].phase, FuturePhase::Pending);
    events.stream_push_bytes(bytes, vec![7, 0]);
    events.stream_close(bytes);
    events.stream_close(bytes);
    assert_eq!(&*received.borrow(), &[128, 255, 7, 0]);
    assert_eq!(*eof.borrow(), 1);
    assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);

    let typed = events.create_stream_of(Some(ValType::I32));
    events.stream_push(typed, Value::I32(257)).unwrap();
    assert_eq!(events.stream_read_items(typed, 1), [Value::I32(257)]);
    assert!(events.set_byte_consumer(typed, future, Box::new(|_| Ok(()))).is_err());
}

#[test]
fn forwarding_preserves_partial_bytes_and_delivers_closed_source_eof_once() {
    let mut events = EventLoop::new();
    let source = events.create_stream();
    let destination = events.create_stream();
    let future = events.create_future();
    events.stream_push_bytes(source, vec![1, 2, 3]);
    assert_eq!(events.stream_read_bytes(source, 1), [1]);
    events.stream_push_bytes(source, vec![4, 5]);
    events.stream_close(source);
    let received = Rc::new(RefCell::new((Vec::new(), 0)));
    let output = received.clone();
    events.set_byte_consumer(destination, future, Box::new(move |chunk| {
        let mut output = output.borrow_mut();
        match chunk {
            Some(chunk) => output.0.extend_from_slice(chunk),
            None => output.1 += 1,
        }
        Ok(())
    })).unwrap();
    events.stream_push_bytes(destination, vec![0]);
    assert!(events.forward_stream(source, destination));
    events.stream_close(destination);
    assert_eq!(received.borrow().0, [0, 2, 3, 4, 5]);
    assert_eq!(received.borrow().1, 1);
    assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
}

#[test]
fn failed_sink_stops_callbacks_and_preserves_its_error_through_close() {
    let mut events = EventLoop::new();
    let stream = events.create_stream();
    let future = events.create_future();
    let calls = Rc::new(RefCell::new(0));
    let observed = calls.clone();
    events.set_byte_consumer(stream, future, Box::new(move |_| {
        *observed.borrow_mut() += 1;
        Err(Value::I32(42))
    })).unwrap();
    events.stream_push_bytes(stream, vec![7]);
    assert!(events.stream_is_eof(stream));
    events.stream_close(stream);
    events.stream_close(stream);
    assert_eq!(*calls.borrow(), 1);
    assert_eq!(events.future_states[&future].value, Some(Value::I32(42)));
}

#[test]
fn forwarding_lends_original_runs_without_reassembly() {
    let mut events = EventLoop::new();
    let source = events.create_stream();
    let destination = events.create_stream();
    let future = events.create_future();
    let first = vec![37; 4096];
    let second = vec![41; 4096];
    let original = [first.as_ptr() as usize, second.as_ptr() as usize];
    events.stream_push_bytes(source, first);
    events.stream_push_bytes(source, second);
    events.stream_close(source);
    let seen = Rc::new(RefCell::new(Vec::new()));
    let captured = seen.clone();
    events.set_byte_consumer(destination, future, Box::new(move |bytes| {
        if let Some(bytes) = bytes {
            captured.borrow_mut().push(bytes.as_ptr() as usize);
        }
        Ok(())
    })).unwrap();
    assert!(events.forward_stream(source, destination));
    assert_eq!(seen.borrow().as_slice(), original);
    assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
}
