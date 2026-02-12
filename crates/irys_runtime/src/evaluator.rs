use crate::environment::Environment;
use crate::value::{RuntimeError, Value};
use irys_parser::Expression;

pub fn evaluate(expr: &Expression, env: &Environment) -> Result<Value, RuntimeError> {
    match expr {
        Expression::IntegerLiteral(i) => Ok(Value::Integer(*i)),
        Expression::DoubleLiteral(d) => Ok(Value::Double(*d)),
        Expression::StringLiteral(s) => Ok(Value::String(s.clone())),
        Expression::BooleanLiteral(b) => Ok(Value::Boolean(*b)),
        Expression::DateLiteral(s) => {
            // Parse the date string from #...# literal and convert to OLE date
            crate::builtins::cdate_fn(&[Value::String(s.clone())])
        }
        Expression::Nothing => Ok(Value::Nothing),

        Expression::Variable(name) => env.get(name.as_str()),

        Expression::MemberAccess(obj, member) => {
            if let Expression::Variable(name) = obj.as_ref() {
                let flat_key = format!("{}.{}", name.as_str(), member.as_str());
                if let Ok(val) = env.get(&flat_key) {
                    return Ok(val);
                }
            }

            let obj_value = evaluate(obj, env)?;
            match obj_value {
                Value::Object(obj_rc) => {
                    obj_rc.borrow().fields.get(member.as_str())
                        .cloned()
                        .ok_or_else(|| RuntimeError::UndefinedVariable(member.as_str().to_string()))
                }
                _ => Err(RuntimeError::TypeError {
                    expected: "Object".to_string(),
                    got: format!("{:?}", obj_value),
                }),
            }
        }

        Expression::Add(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => Ok(Value::Integer(a + b)),
                (Value::Date(d), Value::Double(n)) => Ok(Value::Date(d + n)),
                (Value::Double(n), Value::Date(d)) => Ok(Value::Date(d + n)),
                (Value::Date(d), Value::Integer(n)) => Ok(Value::Date(d + *n as f64)),
                (Value::Integer(n), Value::Date(d)) => Ok(Value::Date(d + *n as f64)),
                (Value::Long(a), Value::Long(b)) => Ok(Value::Long(a + b)),
                (Value::Integer(a), Value::Long(b)) => Ok(Value::Long(*a as i64 + *b)),
                (Value::Long(a), Value::Integer(b)) => Ok(Value::Long(*a + *b as i64)),
                _ => {
                    let a = l.as_double()?;
                    let b = r.as_double()?;
                    Ok(Value::Double(a + b))
                }
            }
        }
        Expression::Subtract(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let res = match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => Ok(Value::Integer(a - b)),
                (Value::Date(d), Value::Double(n)) => Ok(Value::Date(d - n)),
                (Value::Date(d1), Value::Date(d2)) => Ok(Value::Double(d1 - d2)),
                (Value::Date(d), Value::Integer(n)) => Ok(Value::Date(d - *n as f64)),
                (Value::Long(a), Value::Long(b)) => Ok(Value::Long(a - b)),
                (Value::Integer(a), Value::Long(b)) => Ok(Value::Long(*a as i64 - *b)),
                (Value::Long(a), Value::Integer(b)) => Ok(Value::Long(*a - *b as i64)),
                _ => {
                    let a = l.as_double()?;
                    let b = r.as_double()?;
                    Ok(Value::Double(a - b))
                }
            };
            res
        }

        Expression::Multiply(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => Ok(Value::Integer(a * b)),
                (Value::Long(a), Value::Long(b)) => Ok(Value::Long(a * b)),
                (Value::Integer(a), Value::Long(b)) => Ok(Value::Long(*a as i64 * *b)),
                (Value::Long(a), Value::Integer(b)) => Ok(Value::Long(*a * *b as i64)),
                _ => {
                    let a = l.as_double()?;
                    let b = r.as_double()?;
                    Ok(Value::Double(a * b))
                }
            }
        }

        Expression::Divide(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let a = l.as_double()?;
            let b = r.as_double()?;

            if b == 0.0 {
                return Err(RuntimeError::DivisionByZero);
            }

            Ok(Value::Double(a / b))
        }

        Expression::Modulo(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let a = l.as_integer()?;
            let b = r.as_integer()?;

            if b == 0 {
                return Err(RuntimeError::DivisionByZero);
            }

            Ok(Value::Integer(a % b))
        }

        Expression::Concatenate(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            Ok(Value::String(format!("{}{}", l.as_string(), r.as_string())))
        }

        Expression::Equal(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            Ok(Value::Boolean(values_equal(&l, &r)))
        }

        Expression::NotEqual(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            Ok(Value::Boolean(!values_equal(&l, &r)))
        }

        Expression::LessThan(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let result = match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => a < b,
                (Value::String(a), Value::String(b)) => a < b,
                _ => l.as_double()? < r.as_double()?,
            };

            Ok(Value::Boolean(result))
        }

        Expression::LessThanOrEqual(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let result = match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => a <= b,
                (Value::String(a), Value::String(b)) => a <= b,
                _ => l.as_double()? <= r.as_double()?,
            };

            Ok(Value::Boolean(result))
        }

        Expression::GreaterThan(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let result = match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => a > b,
                (Value::String(a), Value::String(b)) => a > b,
                _ => l.as_double()? > r.as_double()?,
            };

            Ok(Value::Boolean(result))
        }

        Expression::GreaterThanOrEqual(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;

            let result = match (&l, &r) {
                (Value::Integer(a), Value::Integer(b)) => a >= b,
                (Value::String(a), Value::String(b)) => a >= b,
                _ => l.as_double()? >= r.as_double()?,
            };

            Ok(Value::Boolean(result))
        }

        Expression::And(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;
            match (&l, &r) {
                (Value::Boolean(a), Value::Boolean(b)) => Ok(Value::Boolean(*a && *b)),
                _ => {
                    let i_l = l.as_long()?;
                    let i_r = r.as_long()?;
                    Ok(Value::Long(i_l & i_r))
                }
            }
        }

        Expression::Or(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;
            match (&l, &r) {
                (Value::Boolean(a), Value::Boolean(b)) => Ok(Value::Boolean(*a || *b)),
                _ => {
                    let i_l = l.as_long()?;
                    let i_r = r.as_long()?;
                    Ok(Value::Long(i_l | i_r))
                }
            }
        }

        Expression::Xor(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;
            match (&l, &r) {
                (Value::Boolean(a), Value::Boolean(b)) => Ok(Value::Boolean(*a ^ *b)),
                _ => {
                    let i_l = l.as_long()?;
                    let i_r = r.as_long()?;
                    Ok(Value::Long(i_l ^ i_r))
                }
            }
        }

        Expression::BitShiftLeft(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;
            let val = l.as_long()?;
            let shift = r.as_integer()?;
            let shift_masked = shift & 63;
            Ok(Value::Long(val << shift_masked))
        }

        Expression::BitShiftRight(left, right) => {
            let l = evaluate(left, env)?;
            let r = evaluate(right, env)?;
            let val = l.as_long()?;
            let shift = r.as_integer()?;
            let shift_masked = shift & 63;
            Ok(Value::Long(val >> shift_masked))
        }

        Expression::Not(operand) => {
            let val = evaluate(operand, env)?;
            match val {
                Value::Boolean(b) => Ok(Value::Boolean(!b)),
                _ => {
                    let i = val.as_long()?;
                    Ok(Value::Long(!i))
                }
            }
        }

        Expression::Negate(operand) => {
            let val = evaluate(operand, env)?;
            match val {
                Value::Integer(i) => Ok(Value::Integer(-i)),
                Value::Double(d) => Ok(Value::Double(-d)),
                _ => {
                    let d = val.as_double()?;
                    Ok(Value::Double(-d))
                }
            }
        }

        Expression::ArrayLiteral(elements) => {
            let vals: Result<Vec<Value>, RuntimeError> = elements
                .iter()
                .map(|e| evaluate(e, env))
                .collect();
            Ok(Value::Array(vals?))
        }

        Expression::ArrayAccess(array, indices) => {
            let arr = env.get(array.as_str())?;
            let index = evaluate(&indices[0], env)?.as_integer()? as usize;
            arr.get_array_element(index)
        }

        Expression::Call(_, _) | Expression::MethodCall(_, _, _) | Expression::New(_, _) | Expression::Me | Expression::WithTarget | Expression::IfExpression(_, _, _) | Expression::AddressOf(_) => {
            // These are handled in the interpreter
            Err(RuntimeError::Custom("Function calls must be evaluated in interpreter context".to_string()))
        }
        Expression::Lambda { .. } => {
            Err(RuntimeError::Custom("Lambdas cannot be evaluated in constant expressions".to_string()))
        }
        Expression::Await(_) => {
            Err(RuntimeError::Custom("Await cannot be evaluated in constant expressions".to_string()))
        }
    }
}

pub fn values_equal(l: &Value, r: &Value) -> bool {
    match (l, r) {
        (Value::Boolean(a), Value::Boolean(b)) => a == b,
        (Value::Integer(a), Value::Integer(b)) => a == b,
        (Value::Long(a), Value::Long(b)) => a == b,
        (Value::Single(a), Value::Single(b)) => a == b,
        (Value::Double(a), Value::Double(b)) => a == b,
        (Value::Date(a), Value::Date(b)) => a == b,
        (Value::String(a), Value::String(b)) => a.eq_ignore_ascii_case(b),
        (Value::Nothing, Value::Nothing) => true,
        (Value::Object(a), Value::Object(b)) => std::rc::Rc::ptr_eq(a, b),
        // Coercion
        (Value::String(s), Value::Nothing) | (Value::Nothing, Value::String(s)) => s.is_empty(),
        (Value::String(s), other) | (other, Value::String(s)) => {
            s.eq_ignore_ascii_case(&other.as_string())
        }
        (Value::Integer(i), Value::Boolean(b)) | (Value::Boolean(b), Value::Integer(i)) => {
            (*i != 0) == *b
        }
        _ => {
            if let (Ok(ld), Ok(rd)) = (l.as_double(), r.as_double()) {
                ld == rd
            } else {
                false
            }
        }
    }
}

pub fn value_in_range(val: &Value, from: &Value, to: &Value) -> bool {
    match (val, from, to) {
        (Value::Integer(v), Value::Integer(f), Value::Integer(t)) => *v >= *f && *v <= *t,
        (Value::String(v), Value::String(f), Value::String(t)) => v >= f && v <= t,
        _ => {
            if let (Ok(v), Ok(f), Ok(t)) = (val.as_double(), from.as_double(), to.as_double()) {
                v >= f && v <= t
            } else {
                false
            }
        }
    }
}

pub fn compare_values(a: &Value, op: &irys_parser::CompOp, b: &Value) -> Result<bool, RuntimeError> {
    use irys_parser::CompOp;

    let result = match op {
        CompOp::Equal => values_equal(a, b),
        CompOp::NotEqual => !values_equal(a, b),
        CompOp::LessThan => match (a, b) {
            (Value::Integer(x), Value::Integer(y)) => x < y,
            (Value::String(x), Value::String(y)) => x < y,
            _ => a.as_double()? < b.as_double()?,
        },
        CompOp::LessThanOrEqual => match (a, b) {
            (Value::Integer(x), Value::Integer(y)) => x <= y,
            (Value::String(x), Value::String(y)) => x <= y,
            _ => a.as_double()? <= b.as_double()?,
        },
        CompOp::GreaterThan => match (a, b) {
            (Value::Integer(x), Value::Integer(y)) => x > y,
            (Value::String(x), Value::String(y)) => x > y,
            _ => a.as_double()? > b.as_double()?,
        },
        CompOp::GreaterThanOrEqual => match (a, b) {
            (Value::Integer(x), Value::Integer(y)) => x >= y,
            (Value::String(x), Value::String(y)) => x >= y,
            _ => a.as_double()? >= b.as_double()?,
        },
    };

    Ok(result)
}
