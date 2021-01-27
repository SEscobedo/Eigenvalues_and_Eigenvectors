Public Class Form1

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        

        If ComboBox1.Text = "2x2" Then
            Dim A(1, 1) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)

            TextBox3.Text = Join(Eigenvalue2x2(A), ", ")

        ElseIf ComboBox1.Text = "3x3" Then
            Dim A(2, 2) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(0, 2) = Val(TextBox8.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)
            A(1, 2) = Val(TextBox9.Text)
            A(2, 0) = Val(TextBox4.Text)
            A(2, 1) = Val(TextBox7.Text)
            A(2, 2) = Val(TextBox10.Text)

            If SymetriMatrix(A) Then

                TextBox3.Text = Join(Eigenvalue3x3(A), ",")
            Else
                ' TextBox3.Text = EigenvaluePO(A, 50)
                TextBox3.Text = "La matriz no es simétrica"
            End If
        ElseIf ComboBox1.Text = "4x4" Then
            Dim A(3, 3) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(0, 2) = Val(TextBox8.Text)
            A(0, 3) = Val(TextBox14.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)
            A(1, 2) = Val(TextBox9.Text)
            A(1, 3) = Val(TextBox13.Text)
            A(2, 0) = Val(TextBox4.Text)
            A(2, 1) = Val(TextBox7.Text)
            A(2, 2) = Val(TextBox10.Text)
            A(2, 3) = Val(TextBox12.Text)
            A(3, 0) = Val(TextBox18.Text)
            A(3, 1) = Val(TextBox17.Text)
            A(3, 2) = Val(TextBox16.Text)
            A(3, 3) = Val(TextBox15.Text)

            TextBox3.Text = EigenvaluePO(A, 50)

        End If
    End Sub
    

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ReDim M(2, 2)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        capturarM()

        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox9.Text = ""
        TextBox4.Text = ""
        TextBox7.Text = ""
        TextBox10.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

        op = "+"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim N(,) As Double
        Dim R(,) As Double

        If ComboBox1.Text = "2x2" Then
            ReDim N(1, 1)
            ReDim R(1, 1)

            N(0, 0) = Val(TextBox1.Text)
            N(0, 1) = Val(TextBox5.Text)
            N(1, 0) = Val(TextBox2.Text)
            N(1, 1) = Val(TextBox6.Text)

        ElseIf ComboBox1.Text = "3x3" Then
            ReDim N(2, 2)
            ReDim R(2, 2)

            N(0, 0) = Val(TextBox1.Text)
            N(0, 1) = Val(TextBox5.Text)
            N(0, 2) = Val(TextBox8.Text)

            N(1, 0) = Val(TextBox2.Text)
            N(1, 1) = Val(TextBox6.Text)
            N(1, 2) = Val(TextBox9.Text)

            N(2, 0) = Val(TextBox4.Text)
            N(2, 1) = Val(TextBox7.Text)
            N(2, 2) = Val(TextBox10.Text)



        ElseIf ComboBox1.Text = "4x4" Then
            ReDim N(3, 3)
            ReDim R(3, 3)

            N(0, 0) = Val(TextBox1.Text)
            N(0, 1) = Val(TextBox5.Text)
            N(0, 2) = Val(TextBox8.Text)
            N(0, 3) = Val(TextBox14.Text)
            N(1, 0) = Val(TextBox2.Text)
            N(1, 1) = Val(TextBox6.Text)
            N(1, 2) = Val(TextBox9.Text)
            N(1, 3) = Val(TextBox13.Text)
            N(2, 0) = Val(TextBox4.Text)
            N(2, 1) = Val(TextBox7.Text)
            N(2, 2) = Val(TextBox10.Text)
            N(2, 3) = Val(TextBox12.Text)
            N(3, 0) = Val(TextBox18.Text)
            N(3, 1) = Val(TextBox17.Text)
            N(3, 2) = Val(TextBox16.Text)
            N(3, 3) = Val(TextBox15.Text)
        End If
        

        If op = "+" Then

            R = SumMatrix(N, M)

        ElseIf op = "-" Then

            R = SumMatrix(M, MatrixEscalar(-1, N))

        ElseIf op = "x" Then

            R = ProdMatrix(M, N)


        End If

      

        If ComboBox1.Text = "2x2" Then
            TextBox1.Text = R(0, 0)
            TextBox5.Text = R(0, 1)
            TextBox2.Text = R(1, 0)
            TextBox6.Text = R(1, 1)
            
        ElseIf ComboBox1.Text = "3x3" Then
            TextBox1.Text = R(0, 0)
            TextBox5.Text = R(0, 1)
            TextBox8.Text = R(0, 2)
            TextBox2.Text = R(1, 0)
            TextBox6.Text = R(1, 1)
            TextBox9.Text = R(1, 2)
            TextBox4.Text = R(2, 0)
            TextBox7.Text = R(2, 1)
            TextBox10.Text = R(2, 2)

        ElseIf ComboBox1.Text = "4x4" Then
            TextBox1.Text = R(0, 0)
            TextBox5.Text = R(0, 1)
            TextBox8.Text = R(0, 2)
            TextBox14.Text = R(0, 3)
            TextBox2.Text = R(1, 0)
            TextBox6.Text = R(1, 1)
            TextBox9.Text = R(1, 2)
            TextBox13.Text = R(1, 3)
            TextBox4.Text = R(2, 0)
            TextBox7.Text = R(2, 1)
            TextBox10.Text = R(2, 2)
            TextBox12.Text = R(2, 3)
            TextBox18.Text = R(3, 0)
            TextBox17.Text = R(3, 1)
            TextBox16.Text = R(3, 2)
            TextBox15.Text = R(3, 3)

        End If

        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        capturarM()

        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox9.Text = ""
        TextBox4.Text = ""
        TextBox7.Text = ""
        TextBox10.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

        op = "-"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        capturarM()

        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox9.Text = ""
        TextBox4.Text = ""
        TextBox7.Text = ""
        TextBox10.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""



        op = "x"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox1.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox9.Text = ""
        TextBox4.Text = ""
        TextBox7.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox3.Text = ""
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

        ReDim M(2, 2)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox1.Text = "1"
        TextBox5.Text = "0"
        TextBox8.Text = "0"
        TextBox2.Text = "0"
        TextBox6.Text = "1"
        TextBox9.Text = "0"
        TextBox4.Text = "0"
        TextBox7.Text = "0"
        TextBox10.Text = "1"
        TextBox12.Text = "0"
        TextBox13.Text = "0"
        TextBox14.Text = "0"
        TextBox15.Text = "1"
        TextBox16.Text = "0"
        TextBox17.Text = "0"
        TextBox18.Text = "0"
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox1.Text = "0"
        TextBox5.Text = "0"
        TextBox8.Text = "0"
        TextBox2.Text = "0"
        TextBox6.Text = "0"
        TextBox9.Text = "0"
        TextBox4.Text = "0"
        TextBox7.Text = "0"
        TextBox10.Text = "0"
        TextBox12.Text = "0"
        TextBox13.Text = "0"
        TextBox14.Text = "0"
        TextBox15.Text = "0"
        TextBox16.Text = "0"
        TextBox17.Text = "0"
        TextBox18.Text = "0"
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim R(3, 3) As Double
        R(0, 0) = Val(TextBox1.Text)
        R(0, 1) = Val(TextBox5.Text)
        R(0, 2) = Val(TextBox8.Text)
        R(0, 3) = Val(TextBox14.Text)

        R(1, 0) = Val(TextBox2.Text)
        R(1, 1) = Val(TextBox6.Text)
        R(1, 2) = Val(TextBox9.Text)
        R(1, 3) = Val(TextBox13.Text)

        R(2, 0) = Val(TextBox4.Text)
        R(2, 1) = Val(TextBox7.Text)
        R(2, 2) = Val(TextBox10.Text)
        R(2, 3) = Val(TextBox12.Text)

        R(3, 0) = Val(TextBox18.Text)
        R(3, 1) = Val(TextBox17.Text)
        R(3, 2) = Val(TextBox16.Text)
        R(3, 3) = Val(TextBox15.Text)


        R = Traspuesta(R)

        TextBox1.Text = R(0, 0)
        TextBox5.Text = R(0, 1)
        TextBox8.Text = R(0, 2)
        TextBox14.Text = R(0, 3)
        TextBox2.Text = R(1, 0)
        TextBox6.Text = R(1, 1)
        TextBox9.Text = R(1, 2)
        TextBox13.Text = R(1, 3)
        TextBox4.Text = R(2, 0)
        TextBox7.Text = R(2, 1)
        TextBox10.Text = R(2, 2)
        TextBox12.Text = R(2, 3)
        TextBox18.Text = R(3, 0)
        TextBox17.Text = R(3, 1)
        TextBox16.Text = R(3, 2)
        TextBox15.Text = R(3, 3)

        TextBox20.Text = ""
        TextBox3.Text = ""

        TextBox21.Text = ""

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        capturarM()


        TextBox11.Text = det(M)

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If ComboBox1.Text = "2x2" Then
            TextBox1.Text = M(0, 0)
            TextBox5.Text = M(0, 1)

            TextBox2.Text = M(1, 0)
            TextBox6.Text = M(1, 1)


        ElseIf ComboBox1.Text = "3x3" Then
            TextBox1.Text = M(0, 0)
            TextBox5.Text = M(0, 1)
            TextBox8.Text = M(0, 2)

            TextBox2.Text = M(1, 0)
            TextBox6.Text = M(1, 1)
            TextBox9.Text = M(1, 2)

            TextBox4.Text = M(2, 0)
            TextBox7.Text = M(2, 1)
            TextBox10.Text = M(2, 2)


        ElseIf ComboBox1.Text = "4x4" Then
            TextBox1.Text = M(0, 0)
            TextBox5.Text = M(0, 1)
            TextBox8.Text = M(0, 2)
            TextBox14.Text = M(0, 3)
            TextBox2.Text = M(1, 0)
            TextBox6.Text = M(1, 1)
            TextBox9.Text = M(1, 2)
            TextBox13.Text = M(1, 3)
            TextBox4.Text = M(2, 0)
            TextBox7.Text = M(2, 1)
            TextBox10.Text = M(2, 2)
            TextBox12.Text = M(2, 3)
            TextBox18.Text = M(3, 0)
            TextBox17.Text = M(3, 1)
            TextBox16.Text = M(3, 2)
            TextBox15.Text = M(3, 3)
        End If


        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "2x2" Then

            TextBox4.Visible = False
            TextBox7.Visible = False
            TextBox8.Visible = False
            TextBox9.Visible = False
            TextBox10.Visible = False
            TextBox12.Visible = False
            TextBox13.Visible = False
            TextBox14.Visible = False
            TextBox15.Visible = False
            TextBox16.Visible = False
            TextBox17.Visible = False
            TextBox18.Visible = False
            ReDim M(1, 1)

        ElseIf ComboBox1.Text = "3x3" Then
            TextBox4.Visible = True
            TextBox7.Visible = True
            TextBox8.Visible = True
            TextBox9.Visible = True
            TextBox10.Visible = True
            TextBox12.Visible = False
            TextBox13.Visible = False
            TextBox14.Visible = False
            TextBox15.Visible = False
            TextBox16.Visible = False
            TextBox17.Visible = False
            TextBox18.Visible = False
            ReDim M(2, 2)

        ElseIf ComboBox1.Text = "4x4" Then
            TextBox4.Visible = True
            TextBox7.Visible = True
            TextBox8.Visible = True
            TextBox9.Visible = True
            TextBox10.Visible = True
            TextBox12.Visible = True
            TextBox13.Visible = True
            TextBox14.Visible = True
            TextBox15.Visible = True
            TextBox16.Visible = True
            TextBox17.Visible = True
            TextBox18.Visible = True
            ReDim M(3, 3)
        End If
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Sub capturarM()


        If ComboBox1.Text = "2x2" Then

            M(0, 0) = Val(TextBox1.Text)
            M(0, 1) = Val(TextBox5.Text)
            M(1, 0) = Val(TextBox2.Text)
            M(1, 1) = Val(TextBox6.Text)

        ElseIf ComboBox1.Text = "3x3" Then
           
            M(0, 0) = Val(TextBox1.Text)
            M(0, 1) = Val(TextBox5.Text)
            M(0, 2) = Val(TextBox8.Text)

            M(1, 0) = Val(TextBox2.Text)
            M(1, 1) = Val(TextBox6.Text)
            M(1, 2) = Val(TextBox9.Text)

            M(2, 0) = Val(TextBox4.Text)
            M(2, 1) = Val(TextBox7.Text)
            M(2, 2) = Val(TextBox10.Text)

            

        ElseIf ComboBox1.Text = "4x4" Then
            M(0, 0) = Val(TextBox1.Text)
            M(0, 1) = Val(TextBox5.Text)
            M(0, 2) = Val(TextBox8.Text)
            M(0, 3) = Val(TextBox14.Text)
            M(1, 0) = Val(TextBox2.Text)
            M(1, 1) = Val(TextBox6.Text)
            M(1, 2) = Val(TextBox9.Text)
            M(1, 3) = Val(TextBox13.Text)
            M(2, 0) = Val(TextBox4.Text)
            M(2, 1) = Val(TextBox7.Text)
            M(2, 2) = Val(TextBox10.Text)
            M(2, 3) = Val(TextBox12.Text)
            M(3, 0) = Val(TextBox18.Text)
            M(3, 1) = Val(TextBox17.Text)
            M(3, 2) = Val(TextBox16.Text)
            M(3, 3) = Val(TextBox15.Text)
        End If

     
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        capturarM()
        Dim R = M

        R = MatrixEscalar(Val(TextBox19.Text), R)

        If ComboBox1.Text = "2x2" Then
            TextBox1.Text = R(0, 0)
            TextBox5.Text = R(0, 1)
            TextBox2.Text = R(1, 0)
            TextBox6.Text = R(1, 1)

        ElseIf ComboBox1.Text = "3x3" Then
            TextBox1.Text = R(0, 0)
            TextBox5.Text = R(0, 1)
            TextBox8.Text = R(0, 2)
            TextBox2.Text = R(1, 0)
            TextBox6.Text = R(1, 1)
            TextBox9.Text = R(1, 2)
            TextBox4.Text = R(2, 0)
            TextBox7.Text = R(2, 1)
            TextBox10.Text = R(2, 2)

        ElseIf ComboBox1.Text = "4x4" Then
            TextBox1.Text = R(0, 0)
            TextBox5.Text = R(0, 1)
            TextBox8.Text = R(0, 2)
            TextBox14.Text = R(0, 3)
            TextBox2.Text = R(1, 0)
            TextBox6.Text = R(1, 1)
            TextBox9.Text = R(1, 2)
            TextBox13.Text = R(1, 3)
            TextBox4.Text = R(2, 0)
            TextBox7.Text = R(2, 1)
            TextBox10.Text = R(2, 2)
            TextBox12.Text = R(2, 3)
            TextBox18.Text = R(3, 0)
            TextBox17.Text = R(3, 1)
            TextBox16.Text = R(3, 2)
            TextBox15.Text = R(3, 3)

        End If


    End Sub

    Private Sub TextBox19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub TextBox20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If ComboBox1.Text = "2x2" Then
            Dim A(1, 1) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)

            TextBox20.Text = EigenvaluePO(A, 50)

        ElseIf ComboBox1.Text = "3x3" Then
            Dim A(2, 2) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(0, 2) = Val(TextBox8.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)
            A(1, 2) = Val(TextBox9.Text)
            A(2, 0) = Val(TextBox4.Text)
            A(2, 1) = Val(TextBox7.Text)
            A(2, 2) = Val(TextBox10.Text)

            TextBox20.Text = EigenvaluePO(A, 50)
            
        ElseIf ComboBox1.Text = "4x4" Then
            Dim A(3, 3) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(0, 2) = Val(TextBox8.Text)
            A(0, 3) = Val(TextBox14.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)
            A(1, 2) = Val(TextBox9.Text)
            A(1, 3) = Val(TextBox13.Text)
            A(2, 0) = Val(TextBox4.Text)
            A(2, 1) = Val(TextBox7.Text)
            A(2, 2) = Val(TextBox10.Text)
            A(2, 3) = Val(TextBox12.Text)
            A(3, 0) = Val(TextBox18.Text)
            A(3, 1) = Val(TextBox17.Text)
            A(3, 2) = Val(TextBox16.Text)
            A(3, 3) = Val(TextBox15.Text)

            TextBox20.Text = EigenvaluePO(A, 50)

        End If

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If ComboBox1.Text = "2x2" Then
            Dim A(1, 1) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)

            If SymetriMatrix(A) Then
                TextBox21.Text = Jacobi(A)
            Else
                TextBox21.Text = "La matriz no es simétrica"
            End If

        ElseIf ComboBox1.Text = "3x3" Then
            Dim A(2, 2) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(0, 2) = Val(TextBox8.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)
            A(1, 2) = Val(TextBox9.Text)
            A(2, 0) = Val(TextBox4.Text)
            A(2, 1) = Val(TextBox7.Text)
            A(2, 2) = Val(TextBox10.Text)

            If SymetriMatrix(A) Then
                TextBox21.Text = Jacobi(A)
            Else
                TextBox21.Text = "La matriz no es simétrica"
            End If

        ElseIf ComboBox1.Text = "4x4" Then
            Dim A(3, 3) As Double

            A(0, 0) = Val(TextBox1.Text)
            A(0, 1) = Val(TextBox5.Text)
            A(0, 2) = Val(TextBox8.Text)
            A(0, 3) = Val(TextBox14.Text)
            A(1, 0) = Val(TextBox2.Text)
            A(1, 1) = Val(TextBox6.Text)
            A(1, 2) = Val(TextBox9.Text)
            A(1, 3) = Val(TextBox13.Text)
            A(2, 0) = Val(TextBox4.Text)
            A(2, 1) = Val(TextBox7.Text)
            A(2, 2) = Val(TextBox10.Text)
            A(2, 3) = Val(TextBox12.Text)
            A(3, 0) = Val(TextBox18.Text)
            A(3, 1) = Val(TextBox17.Text)
            A(3, 2) = Val(TextBox16.Text)
            A(3, 3) = Val(TextBox15.Text)

            If SymetriMatrix(A) Then
                TextBox21.Text = Jacobi(A)
            Else
                TextBox21.Text = "La matriz no es simétrica"
            End If

        End If
    End Sub

    Private Sub TextBox21_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox18.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox17_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox17.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox16_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox16.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub TextBox15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox15.TextChanged
        TextBox20.Text = ""
        TextBox3.Text = ""
        TextBox11.Text = ""
        TextBox21.Text = ""
    End Sub
End Class
