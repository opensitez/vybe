use crate::value::{RuntimeError, Value};

/// Pmt(rate, nper, pv[, fv[, type]]) - Calculate payment for a loan
pub fn pmt_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(RuntimeError::Custom("Pmt requires 3 to 5 arguments".to_string()));
    }
    
    let rate = args[0].as_double()?;
    let nper = args[1].as_double()?;
    let pv = args[2].as_double()?;
    let fv = if args.len() >= 4 { args[3].as_double()? } else { 0.0 };
    let pmt_type = if args.len() == 5 { args[4].as_integer()? } else { 0 };
    
    let payment = if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let pvif = (1.0 + rate).powf(nper);
        let payment = -(rate * (pv * pvif + fv)) / (pvif - 1.0);
        
        if pmt_type == 1 {
            payment / (1.0 + rate)
        } else {
            payment
        }
    };
    
    Ok(Value::Double(payment))
}

/// FV(rate, nper, pmt[, pv[, type]]) - Calculate future value of an investment
pub fn fv_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(RuntimeError::Custom("FV requires 3 to 5 arguments".to_string()));
    }
    
    let rate = args[0].as_double()?;
    let nper = args[1].as_double()?;
    let pmt = args[2].as_double()?;
    let pv = if args.len() >= 4 { args[3].as_double()? } else { 0.0 };
    let pmt_type = if args.len() == 5 { args[4].as_integer()? } else { 0 };
    
    let fv = if rate == 0.0 {
        -(pv + pmt * nper)
    } else {
        let pvif = (1.0 + rate).powf(nper);
        let fvifa = ((1.0 + rate).powf(nper) - 1.0) / rate;
        
        let mut future_value = -(pv * pvif + pmt * fvifa);
        
        if pmt_type == 1 {
            future_value /= 1.0 + rate;
        }
        
        future_value
    };
    
    Ok(Value::Double(fv))
}

/// PV(rate, nper, pmt[, fv[, type]]) - Calculate present value of an investment
pub fn pv_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(RuntimeError::Custom("PV requires 3 to 5 arguments".to_string()));
    }
    
    let rate = args[0].as_double()?;
    let nper = args[1].as_double()?;
    let pmt = args[2].as_double()?;
    let fv = if args.len() >= 4 { args[3].as_double()? } else { 0.0 };
    let pmt_type = if args.len() == 5 { args[4].as_integer()? } else { 0 };
    
    let present_value = if rate == 0.0 {
        -(fv + pmt * nper)
    } else {
        let pvif = (1.0 + rate).powf(-nper);
        let fvifa = ((1.0 + rate).powf(nper) - 1.0) / (rate * (1.0 + rate).powf(nper));
        
        let mut pv = -(fv * pvif + pmt * fvifa);
        
        if pmt_type == 1 {
            pv /= 1.0 + rate;
        }
        
        pv
    };
    
    Ok(Value::Double(present_value))
}

/// NPer(rate, pmt, pv[, fv[, type]]) - Calculate number of periods for an investment
pub fn nper_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(RuntimeError::Custom("NPer requires 3 to 5 arguments".to_string()));
    }
    
    let rate = args[0].as_double()?;
    let pmt = args[1].as_double()?;
    let pv = args[2].as_double()?;
    let fv = if args.len() >= 4 { args[3].as_double()? } else { 0.0 };
    let pmt_type = if args.len() == 5 { args[4].as_integer()? } else { 0 };
    
    if rate == 0.0 {
        return Ok(Value::Double(-(pv + fv) / pmt));
    }
    
    let adjusted_pmt = if pmt_type == 1 {
        pmt * (1.0 + rate)
    } else {
        pmt
    };
    
    let num = adjusted_pmt - fv * rate;
    let denom = adjusted_pmt + pv * rate;
    
    if num <= 0.0 || denom <= 0.0 {
        return Err(RuntimeError::Custom("NPer: invalid parameters".to_string()));
    }
    
    let nper = (num / denom).ln() / (1.0 + rate).ln();
    Ok(Value::Double(nper))
}

/// Rate(nper, pmt, pv[, fv[, type[, guess]]]) - Calculate interest rate
pub fn rate_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 || args.len() > 6 {
        return Err(RuntimeError::Custom("Rate requires 3 to 6 arguments".to_string()));
    }
    
    let nper = args[0].as_double()?;
    let pmt = args[1].as_double()?;
    let pv = args[2].as_double()?;
    let fv = if args.len() >= 4 { args[3].as_double()? } else { 0.0 };
    let pmt_type = if args.len() >= 5 { args[4].as_integer()? } else { 0 };
    let mut guess = if args.len() == 6 { args[5].as_double()? } else { 0.1 };
    
    // Use Newton-Raphson method to find rate
    let max_iterations = 100;
    let precision = 0.00000001;
    
    for _ in 0..max_iterations {
        let f = calculate_rate_function(nper, pmt, pv, fv, pmt_type, guess);
        let df = calculate_rate_derivative(nper, pmt, pv, fv, pmt_type, guess);
        
        if df.abs() < precision {
            break;
        }
        
        let new_guess = guess - f / df;
        
        if (new_guess - guess).abs() < precision {
            return Ok(Value::Double(new_guess));
        }
        
        guess = new_guess;
    }
    
    Ok(Value::Double(guess))
}

fn calculate_rate_function(nper: f64, pmt: f64, pv: f64, fv: f64, pmt_type: i32, rate: f64) -> f64 {
    if rate == 0.0 {
        return pv + pmt * nper + fv;
    }
    
    let pvif = (1.0 + rate).powf(nper);
    let adjusted_pmt = if pmt_type == 1 { pmt * (1.0 + rate) } else { pmt };
    
    pv * pvif + adjusted_pmt * (pvif - 1.0) / rate + fv
}

fn calculate_rate_derivative(nper: f64, pmt: f64, pv: f64, fv: f64, pmt_type: i32, rate: f64) -> f64 {
    let delta = 0.0001;
    let f1 = calculate_rate_function(nper, pmt, pv, fv, pmt_type, rate + delta);
    let f0 = calculate_rate_function(nper, pmt, pv, fv, pmt_type, rate);
    (f1 - f0) / delta
}

/// IPmt(rate, per, nper, pv[, fv[, type]]) - Calculate interest payment for a period
pub fn ipmt_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 4 || args.len() > 6 {
        return Err(RuntimeError::Custom("IPmt requires 4 to 6 arguments".to_string()));
    }
    
    let rate = args[0].as_double()?;
    let per = args[1].as_double()?;
    let nper = args[2].as_double()?;
    let pv = args[3].as_double()?;
    let fv = if args.len() >= 5 { args[4].as_double()? } else { 0.0 };
    let pmt_type = if args.len() == 6 { args[5].as_integer()? } else { 0 };
    
    // Calculate payment first
    let pmt_args = vec![
        Value::Double(rate),
        Value::Double(nper),
        Value::Double(pv),
        Value::Double(fv),
        Value::Integer(pmt_type),
    ];
    let payment = pmt_fn(&pmt_args)?.as_double()?;
    
    // Calculate principal at start of period
    let fv_args = vec![
        Value::Double(rate),
        Value::Double(per - 1.0),
        Value::Double(payment),
        Value::Double(pv),
        Value::Integer(pmt_type),
    ];
    let principal = -fv_fn(&fv_args)?.as_double()?;
    
    let interest = if pmt_type == 1 && per == 1.0 {
        0.0
    } else {
        principal * rate
    };
    
    Ok(Value::Double(interest))
}

/// PPmt(rate, per, nper, pv[, fv[, type]]) - Calculate principal payment for a period
pub fn ppmt_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 4 || args.len() > 6 {
        return Err(RuntimeError::Custom("PPmt requires 4 to 6 arguments".to_string()));
    }
    
    // Calculate total payment
    let pmt_args = vec![
        args[0].clone(),
        args[2].clone(),
        args[3].clone(),
        if args.len() >= 5 { args[4].clone() } else { Value::Double(0.0) },
        if args.len() == 6 { args[5].clone() } else { Value::Integer(0) },
    ];
    let payment = pmt_fn(&pmt_args)?.as_double()?;
    
    // Calculate interest payment
    let interest = ipmt_fn(args)?.as_double()?;
    
    // Principal = Payment - Interest
    let principal = payment - interest;
    
    Ok(Value::Double(principal))
}

/// DDB(cost, salvage, life, period[, factor]) - Calculate depreciation (double-declining balance)
pub fn ddb_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 4 || args.len() > 5 {
        return Err(RuntimeError::Custom("DDB requires 4 or 5 arguments".to_string()));
    }
    
    let cost = args[0].as_double()?;
    let salvage = args[1].as_double()?;
    let life = args[2].as_double()?;
    let period = args[3].as_double()?;
    let factor = if args.len() == 5 { args[4].as_double()? } else { 2.0 };
    
    if life == 0.0 {
        return Err(RuntimeError::Custom("DDB: life cannot be zero".to_string()));
    }
    
    let rate = factor / life;
    let mut book_value = cost;
    let mut depreciation = 0.0;
    
    for _p in 1..=(period as i32) {
        let current_depreciation = book_value * rate;
        let max_depreciation = book_value - salvage;
        
        depreciation = if current_depreciation > max_depreciation {
            max_depreciation
        } else {
            current_depreciation
        };
        
        book_value -= depreciation;
        
        if book_value < salvage {
            break;
        }
    }
    
    Ok(Value::Double(depreciation))
}

/// SLN(cost, salvage, life) - Calculate straight-line depreciation
pub fn sln_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("SLN requires 3 arguments".to_string()));
    }
    
    let cost = args[0].as_double()?;
    let salvage = args[1].as_double()?;
    let life = args[2].as_double()?;
    
    if life == 0.0 {
        return Err(RuntimeError::Custom("SLN: life cannot be zero".to_string()));
    }
    
    let depreciation = (cost - salvage) / life;
    Ok(Value::Double(depreciation))
}

/// SYD(cost, salvage, life, period) - Calculate sum-of-years depreciation
pub fn syd_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 4 {
        return Err(RuntimeError::Custom("SYD requires 4 arguments".to_string()));
    }
    
    let cost = args[0].as_double()?;
    let salvage = args[1].as_double()?;
    let life = args[2].as_double()?;
    let period = args[3].as_double()?;
    
    if life == 0.0 {
        return Err(RuntimeError::Custom("SYD: life cannot be zero".to_string()));
    }
    
    let sum_of_years = (life * (life + 1.0)) / 2.0;
    let depreciation = (cost - salvage) * (life - period + 1.0) / sum_of_years;
    
    Ok(Value::Double(depreciation))
}
