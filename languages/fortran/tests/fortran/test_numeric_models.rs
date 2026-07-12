use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    nm_01,
    "program p
print *, digits(1.0)
end program p
"
);
c!(
    nm_02,
    "program p
print *, epsilon(1.0)
end program p
"
);
c!(
    nm_03,
    "program p
print *, huge(1)
end program p
"
);
c!(
    nm_04,
    "program p
print *, tiny(1.0)
end program p
"
);
c!(
    nm_05,
    "program p
print *, radix(1.0)
end program p
"
);
c!(
    nm_06,
    "program p
print *, range(1.0)
end program p
"
);
c!(
    nm_07,
    "program p
print *, precision(1.0)
end program p
"
);
c!(
    nm_08,
    "program p
print *, spacing(1.0)
end program p
"
);
c!(
    nm_09,
    "program p
print *, nearest(1.0,1.0)
end program p
"
);
c!(
    nm_10,
    "program p
print *, rrspacing(1.0)
end program p
"
);
c!(
    nm_11,
    "program p
print *, scale(1.0,2)
end program p
"
);
c!(
    nm_12,
    "program p
print *, selected_int_kind(9)
end program p
"
);
c!(
    nm_13,
    "program p
print *, selected_real_kind(6)
end program p
"
);
c!(
    nm_14,
    "program p
print *, kind(1.0)
end program p
"
);
c!(
    nm_15,
    "program p
print *, int(1.5)
end program p
"
);
c!(
    nm_16,
    "program p
print *, real(1)
end program p
"
);
c!(
    nm_17,
    "program p
print *, nint(1.6)
end program p
"
);
c!(
    nm_18,
    "program p
print *, floor(1.6)
end program p
"
);
c!(
    nm_19,
    "program p
print *, ceiling(1.2)
end program p
"
);
c!(
    nm_20,
    "program p
print *, sign(2.0,-1.0)
end program p
"
);
