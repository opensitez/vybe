#[path = "fortran/helpers.rs"]
mod helpers;

#[test]
fn probe_format_outputs() {
    let cases = [
        ("I0", "program t\nprint '(I0)', 7\nend program t\n", "7"),
        ("I5", "program t\nprint '(I5)', 42\nend program t\n", "42"),
        ("I5.3", "program t\nprint '(I5.3)', 7\nend program t\n", "7"),
        ("B8", "program t\nprint '(B8)', 15\nend program t\n", "15"),
        ("O4", "program t\nprint '(O4)', 10\nend program t\n", "10"),
        ("Z4", "program t\nprint '(Z4)', 255\nend program t\n", "255"),
        (
            "SP F6.2",
            "program t\nprint '(SP,F6.2)', 1.5\nend program t\n",
            "1.5",
        ),
        (
            "SS F6.2 pos",
            "program t\nprint '(SS,F6.2)', 2.5\nend program t\n",
            "2.50",
        ),
        (
            "BN I4",
            "program t\nprint '(BN,I4)', 7\nend program t\n",
            "7",
        ),
        (
            "EN12.3",
            "program t\nprint '(EN12.3)', 1.23\nend program t\n",
            "1.230e+0",
        ),
        (
            "ES10.3",
            "program t\nprint '(ES10.3)', 0.25\nend program t\n",
            "2.500e-1",
        ),
        (
            "G12.4",
            "program t\nprint '(G12.4)', 3.14\nend program t\n",
            "3.1400e+0",
        ),
        (
            "T5 I0",
            "program t\nprint '(T5,I0)', 9\nend program t\n",
            "9",
        ),
        (
            "TL3 I0",
            "program t\nprint '(TL3,I0)', 9\nend program t\n",
            "9",
        ),
        (
            "TR2 I0",
            "program t\nprint '(TR2,I0)', 9\nend program t\n",
            "9",
        ),
        (
            "dollar",
            "program t\nprint '($,I0)', 5\nend program t\n",
            "5",
        ),
        (
            "colon I0",
            "program t\nprint '(I0,:)', 5\nend program t\n",
            "5",
        ),
        (
            "2(I0)",
            "program t\nprint '(2(I0))', 3, 4\nend program t\n",
            "34",
        ),
        (
            "internal I0",
            "program t\ncharacter(len=4) :: buf\nwrite(buf, '(I0)') 15\nprint *, trim(buf)\nend program t\n",
            "15",
        ),
        (
            "internal I4 read",
            "program t\ncharacter(len=4) :: buf = '  7'\ninteger :: n\nread(buf, '(I4)') n\nprint *, n\nend program t\n",
            "7",
        ),
        (
            "A8",
            "program t\nprint '(A8)', 'Fortran'\nend program t\n",
            "Fortran ",
        ),
        (
            "L1 T",
            "program t\nprint '(L1)', .true.\nend program t\n",
            "true",
        ),
        (
            "3X A",
            "program t\nprint '(A,3X,A)', 'L','R'\nend program t\n",
            "L   R",
        ),
        (
            "slash",
            "program t\nprint '(I0,/,I0)', 1, 2\nend program t\n",
            "1\n2",
        ),
    ];
    for (name, src, _expected) in cases {
        let actual = helpers::run_prints(src);
        eprintln!("{name}: {actual:?}");
    }
}
