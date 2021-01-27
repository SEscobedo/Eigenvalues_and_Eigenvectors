Module Module1



    Public M(,) As Double
    Public op As String
    Dim S(,), e1(), E(,) As Double
    Dim changed() As Boolean
    Dim state As Long


    Function ProdMatrix(ByVal A(,) As Double, ByVal B(,) As Double) As Double(,)
        Dim i, j, k As Long
        Dim res(,) As Double
        ReDim res(UBound(B, 1), UBound(A, 2))

        For i = 0 To UBound(res, 2) 'fila
            For j = 0 To UBound(res, 1) 'columna
                For k = 0 To UBound(A, 1)
                    res(j, i) = res(j, i) + A(k, i) * B(j, k)
                Next
            Next
        Next
        ProdMatrix = res
    End Function
    Function SumMatrix(ByVal A(,) As Double, ByVal B(,) As Double) As Double(,)
        Dim i, j As Long
        Dim res(,) As Double
        ReDim res(UBound(A, 1), UBound(B, 2))

        For i = 0 To UBound(A, 1)
            For j = 0 To UBound(A, 2)
                res(i, j) = A(i, j) + B(i, j)
            Next
        Next

        SumMatrix = res
    End Function

    Function MatrixEscalar(ByVal c As Double, ByVal A(,) As Double) As Double(,)
        Dim i, j As Long

        For i = 0 To UBound(A, 1)
            For j = 0 To UBound(A, 2)
                A(i, j) = A(i, j) * c
            Next
        Next

        MatrixEscalar = A
    End Function


    

    'determina el tipo de matriz.
    Function Tipo(ByVal A(,) As Double) As String

    End Function

    'autovalores de una matriz triangular
    Function EigenvalueTriang(ByVal A(,) As Double) As String()
        'los autovalores son sus elementos diagonales.
        Dim res() As String
        ReDim res(UBound(A, 1))
        Dim i As Long
        For i = 0 To UBound(A, 1)
            res(i) = Str(A(i, i))
        Next
        EigenvalueTriang = res
    End Function
    Function Eigenvalue2x2(ByVal A(,) As Double) As String()

        Dim res(1) As String

        res(0) = 1 / 2 * (A(0, 0) + A(1, 1) - Math.Sqrt(A(0, 0) ^ 2 + 4 * A(0, 1) * A(1, 0) - 2 * A(0, 0) * A(1, 1) + A(1, 1) ^ 2))
        res(1) = 1 / 2 * (A(0, 0) + A(1, 1) + Math.Sqrt(A(0, 0) ^ 2 + 4 * A(0, 1) * A(1, 0) - 2 * A(0, 0) * A(1, 1) + A(1, 1) ^ 2))

        Eigenvalue2x2 = res

    End Function
    Function Eigenvalue3x3(ByVal A(,) As Double) As String()
        '    % Given an real symmetric 3x3 matrix A, compute the eigenvalues
        Dim p, q, r, phi, eig1, eig2, eig3, B(,), I(2, 2) As Double

        I(0, 0) = 1
        I(1, 1) = 1
        I(2, 2) = 1
        I(0, 1) = 0
        I(0, 2) = 0
        I(1, 0) = 0
        I(1, 2) = 0
        I(2, 0) = 0
        I(2, 1) = 0

        p = A(0, 1) ^ 2 + A(0, 2) ^ 2 + A(1, 2) ^ 2

        If (p = 0) Then
            '% A is diagonal.
            eig1 = A(0, 0)
            eig2 = A(1, 1)
            eig3 = A(2, 2)
        Else
            q = (A(0, 0) + A(1, 1) + A(2, 2)) / 3
            p = (A(0, 0) - q) ^ 2 + (A(1, 1) - q) ^ 2 + (A(2, 2) - q) ^ 2 + 2 * p
            p = Math.Sqrt(p / 6)
            B = MatrixEscalar(1 / p, SumMatrix(A, MatrixEscalar(-q, I)))     ' % I is the identity matrix
            r = det(B) / 2

            '% In exact arithmetic for a symmetric matrix  -1 <= r <= 1
            '% but computation error can leave it slightly outside this range.
            If (r <= -1) Then
                phi = Math.PI / 3
            ElseIf (r >= 1) Then
                phi = 0
            Else
                phi = Math.Acos(r) / 3
            End If

            '   % the eigenvalues satisfy eig3 <= eig2 <= eig1
            eig1 = q + 2 * p * Math.Cos(phi)
            eig3 = q + 2 * p * Math.Cos(phi + Math.PI * (2 / 3))
            eig2 = 3 * q - eig1 - eig3   '  % since trace(A) = eig1 + eig2 + eig3
        End If
        Dim res(2) As String
        res(0) = Str(Math.Round(eig1, 4))
        res(1) = Str(Math.Round(eig2, 4))
        res(2) = Str(Math.Round(eig3, 4))


        Eigenvalue3x3 = res

    End Function

    Function det(ByVal A(,) As Double) As Double

        If UBound(A, 1) = 1 Then
            det = -A(1, 0) * A(0, 1) + A(0, 0) * A(1, 1)
        ElseIf UBound(A, 1) = 2 Then
            det = -A(2, 0) * A(1, 1) * A(0, 2) + A(1, 0) * A(2, 1) * A(0, 2) + A(2, 0) * A(0, 1) * A(1, 2) - A(0, 0) * A(2, 1) * A(1, 2) - A(1, 0) * A(0, 1) * A(2, 2) + A(0, 0) * A(1, 1) * A(2, 2)
        ElseIf UBound(A, 1) = 3 Then
            det = A(0, 3) * A(1, 2) * A(2, 1) * A(3, 0) - A(0, 2) * A(1, 3) * A(2, 1) * A(3, 0) - A(0, 3) * A(1, 1) * A(2, 2) * A(3, 0) +
A(0, 1) * A(1, 3) * A(2, 2) * A(3, 0) + A(0, 2) * A(1, 1) * A(2, 3) * A(3, 0) - A(0, 1) * A(1, 2) * A(2, 3) * A(3, 0) -
A(0, 3) * A(1, 2) * A(2, 0) * A(3, 1) + A(0, 2) * A(1, 3) * A(2, 0) * A(3, 1) + A(0, 3) * A(1, 0) * A(2, 2) * A(3, 1) -
A(0, 0) * A(1, 3) * A(2, 2) * A(3, 1) - A(0, 2) * A(1, 0) * A(2, 3) * A(3, 1) + A(0, 0) * A(1, 2) * A(2, 3) * A(3, 1) +
A(0, 3) * A(1, 1) * A(2, 0) * A(3, 2) - A(0, 1) * A(1, 3) * A(2, 0) * A(3, 2) - A(0, 3) * A(1, 0) * A(2, 1) * A(3, 2) +
A(0, 0) * A(1, 3) * A(2, 1) * A(3, 2) + A(0, 1) * A(1, 0) * A(2, 3) * A(3, 2) - A(0, 0) * A(1, 1) * A(2, 3) * A(3, 2) -
A(0, 2) * A(1, 1) * A(2, 0) * A(3, 3) + A(0, 1) * A(1, 2) * A(2, 0) * A(3, 3) + A(0, 2) * A(1, 0) * A(2, 1) * A(3, 3) -
A(0, 0) * A(1, 2) * A(2, 1) * A(3, 3) - A(0, 1) * A(1, 0) * A(2, 2) * A(3, 3) + A(0, 0) * A(1, 1) * A(2, 2) * A(3, 3)

        End If
    End Function

    Function Traspuesta(ByVal A(,) As Double) As Double(,)
        Dim i, j As Long
        Dim R(,) As Double
        ReDim R(UBound(A, 2), UBound(A, 1))
        For i = 0 To UBound(A, 1)
            For j = 0 To UBound(A, 2)
                R(j, i) = A(i, j)
            Next
        Next
        Traspuesta = R
    End Function

    Function SymetriMatrix(ByVal A(,) As Double) As Boolean
        Dim i, j As Long
        Dim res As Boolean

        For i = 0 To UBound(A, 1)
            For j = 0 To UBound(A, 2)
                If A(i, j) = A(j, i) Then
                    res = True
                Else
                    res = False
                    Exit For
                End If
            Next
            If res = False Then Exit For
        Next

        SymetriMatrix = res

        'If Traspuesta(A) = A Then
        '    SymetriMatrix = True
        'Else
        '    SymetriMatrix = False
        'End If

    End Function
    
    'Function EigenvalueIN(ByVal A(,) As Double) As String() 'método de las potencias inversas con translación
    '    'primera(aproximación)
    '    Dim ei() As Double
    '    Dim X(,) As Double
    '    Dim alfa As Double
    '    ReDim ei(UBound(A, 1))
    '    ReDim X(0, UBound(A, 1))
    '    Dim i As Long

    '    For i = 0 To UBound(X, 2)
    '        X(0, i) = 1
    '    Next

    '    'asignar un valor aproximativo a cada autovalor

    '    For i = 0 To UBound(ei)
    '        ei(i) = 2
    '        alfa = ei(i) * 0.05 + ei(i)


    '    Next

    'End Function

    Function EigenvaluePO(ByVal A(,) As Double, ByVal n As Long) As String 'método de las potencias 
        'primera(aproximación)
        Dim ei() As Double
        Dim X(,) As Double
        Dim res() As String
        Dim c As Double
        ReDim ei(UBound(A, 1))
        ReDim X(0, UBound(A, 1))
        ReDim res(UBound(A, 1))
        Dim i, j As Long

        For i = 0 To UBound(X, 2)
            X(0, i) = 1
        Next


        'n = número de iteraciones
        For i = 0 To n
            X = ProdMatrix(A, X)
            For j = 0 To UBound(X, 2)
                c = 0
                c = Math.Max(Math.Abs(c), Math.Abs(X(0, j)))
                c = Math.Sign(X(0, j))
            Next
            If Not c = 0 Then X = MatrixEscalar(1 / c, X)
        Next

        For i = 0 To UBound(X, 2)
            res(i) = Str(Math.Round(X(0, i), 4))
        Next

        EigenvaluePO = "Autovector dominante: (" & Join(res, ",") & "); autovalor dominante: " & Str(Math.Round(c, 4))
    End Function

    Function Jacobi(ByVal A(,) As Double) As String

        'método de Jacobi para el cálculo de autovalores y autovectores

        'procedure jacobi(S ∈ Rn×n; out e ∈ Rn; out E ∈ Rn×n)

        Dim i, k, l, m1, n As Long
        Dim s1, c, t, p, y As Double
        Dim ind() As Long
        n = UBound(A, 1) + 1
        ReDim ind(n)
        ReDim e1(n)
        ReDim E(n, n)
        ReDim changed(n)
        ReDim ind(n)
        ReDim S(n, n)

        ',,,,,,,,,,,,,,,,,,,,,,,
        Dim u, v As Long
        For u = 1 To UBound(A, 1) + 1
            For v = 1 To UBound(A, 2) + 1
                S(u, v) = A(u - 1, v - 1)
            Next
        Next
        ',,,,,,,,,,,,,,,,,,,,,,,,

        '! init e, E, and arrays ind, changed
        'incializar E como matriz identidad

        For u = 1 To n
            For v = 1 To n
                If u = v Then
                    E(u, v) = 1
                Else
                    E(u, v) = 0
                End If
            Next
        Next

        state = n

        For k = 1 To n
            ind(k) = maxind(k, n)
            e1(k) = S(k, k)
            changed(k) = True
        Next
        'Do While Not state = 0 '! next rotation
        Dim iterations = 50 '<--- número de iteraciones a realizar
        For u = 0 To iterations
            m1 = 1 '! find index (k,l) of pivot p
            For k = 2 To n - 1
                If Math.Abs(S(k, ind(k))) > Math.Abs(S(m1, ind(m1))) Then
                    m1 = k
                End If
            Next k
            k = m1
            l = ind(m1)
            p = S(k, l)
            '  ! calculate c = cos φ, s = sin φ
            y = (e1(l) - e1(k)) / 2
            t = Math.Abs(y) + Math.Sqrt(p ^ 2 + y ^ 2)
            s1 = Math.Sqrt(p ^ 2 + t ^ 2)
            c = t / s1
            s1 = p / s1
            'If t = 0 Then
            '    state = 0 'por mientras ¿qué mas se puede hacer?
            'Else
            t = (p ^ 2) / t
            'End If

            If y < 0 Then
                s1 = -s1
                t = -t
            End If

            S(k, l) = 0.0
            update(k, -t)
            update(l, t)
            '  ! rotate rows and columns k and l
            For i = 1 To k - 1
                rotate(i, k, i, l, c, s1)
            Next
            For i = k + 1 To l - 1
                rotate(k, i, i, l, c, s1)
            Next
            For i = l + 1 To n
                rotate(k, i, l, i, c, s1)
            Next
            '                                  !rotate(eigenvectors)
            For i = 1 To n
                Dim A1(0, 1), B1(1, 1) As Double
                A1(0, 0) = E(k, i)
                A1(0, 1) = E(l, i)

                B1(0, 0) = c
                B1(0, 1) = s1
                B1(1, 0) = -s1
                B1(1, 1) = c

                A1 = ProdMatrix(B1, A1)
                '    ┌   ┐    ┌     ┐┌   ┐
                '    │Eki│    │c  −s││Eki│
                '    │   │ := │     ││   │
                '    │Eli│    │s   c││Eli│
                '    └   ┘    └     ┘└   ┘

                E(k, i) = A1(0, 0)
                E(l, i) = A1(0, 1)


            Next
            '  ! rows k, l have changed, update rows indk, indl
            ind(k) = maxind(k, n)
            ind(l) = maxind(l, n)
            'Loop
        Next

        Dim res(n) As String
        For u = 1 To n
            res(u) = Str(Math.Round(e1(u), 4))
        Next
        Dim st, auto As String
        auto = ""
        st = ""

        For u = 1 To n
            For v = 1 To n
                If v = 1 Then
                    auto = Math.Round(E(u, v), 4)
                Else
                    auto = auto & ", " & Math.Round(E(u, v), 4)
                End If

            Next
            st = st & "λ" & u & " = " & res(u) & "  [" & auto & "]'" & vbCrLf
            auto = ""
        Next

        Jacobi = st

    End Function

    Function maxind(ByVal k As Long, ByVal n As Long) As Long ' index of largest off-diagonal element in row k
        Dim res As Long
        res = k + 1
        For i = k + 2 To n
            If Math.Abs(S(k, i)) > Math.Abs(S(k, res)) Then
                res = i
            End If
        Next
        maxind = res
    End Function

    Sub rotate(ByVal k As Long, ByVal l As Long, ByVal i As Long, ByVal j As Long, ByRef c As Double, ByRef s1 As Double) '! perform rotation of Sij, Skl
        Dim A1(0, 1), B1(1, 1) As Double
        A1(0, 0) = S(k, l)
        A1(0, 1) = S(i, j)

        B1(0, 0) = c
        B1(0, 1) = s1
        B1(1, 0) = -s1
        B1(1, 1) = c

        A1 = ProdMatrix(B1, A1)
        '  ┌   ┐    ┌     ┐┌   ┐
        '  │Skl│    │c  −s││Skl│
        '  │   │ := │     ││   │
        '  │Sij│    │s   c││Sij│
        '  └   ┘    └     ┘└   ┘

        S(k, l) = A1(0, 0)
        S(i, j) = A1(0, 1)

    End Sub

    Sub update(ByVal k As Long, ByVal t As Double) '! update ek and its status
        Dim y As Double
        y = e1(k)
        e1(k) = y + t

        If changed(k) And y = e1(k) Then
            changed(k) = False
            state = state - 1

        ElseIf (Not changed(k)) And (Not y = e1(k)) Then
            changed(k) = True
            state = state + 1
        End If
    End Sub


End Module
