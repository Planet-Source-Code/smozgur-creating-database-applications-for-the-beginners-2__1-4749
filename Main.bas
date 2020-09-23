Attribute VB_Name = "DBModule"
Global DBMain As Database
Global RecSet As Recordset

Public Sub CreateNewDB()
    'You've already got this procedure from the first code
End Sub

Public Sub Main()
    Stop
    'Please ReadThis

'!!!WARNING!!!
'First you have to add Microsoft DAO Object Library
'using by References Dialog Box
'If not VB doesnot know the database elements
    
    
    'This is a learning program for beginners
    'Please dont response if you have advanced skill in DB programming
    
    '** QueryData queries your database and little SQL-Now
    
    'If you decide to follow my codes i am going to send
    'them regularly and you will learn everything about
    'database programming.
    
    'You can test the codes first running project
    'putting the breakpoints to the header of the procedures.
    'To run the procedure write the name of the Sub into the
    'immediate window then press Enter.
    'It is the best way to understand the whole code.
    
    'Today's Tip
    'There is a sample filled database which as an API dictionary
    'We will use it today as sample
    'And we have a form as an interface
    'Enjoy it!
    Load Form1
    Form1.Show
End Sub


Public Sub QueryData(reqText As String)
    'first you need to learn little SQL
    'SQL is Structured Query Language
    'it is like this:
    
    'SELECT some fields FROM databases WHERE somehing you choose
    
    'Of course it is not real SQL exactly
    'But you have to use this first time
    
    'I always create my database queries with SQL
    'It is more fast and easy
    'I will teach you how you use other way to querying
    'But this time SQL is the better way
    
    'Let's say the user first hit 'w' key and continue
    'But we run this code from the Change event of box
    'So this code will query the database for each key
    'This means for the first key it bring to me
    'the words begin with 'w'
    
    'Then for the second word lets say 'wo'
    'it will bring to me the words begin with 'wo'
    'And continue...
       
    'The reqText is the word which looking for
    'Then it is not important how many characters if it is
    
 
    'Dimensioning string
Dim SQLText As String
    'Create our SQL Text
    SQLText = "SELECT * FROM DBWords WHERE dbWord LIKE '" & reqText & "*';" 'WHERE Left(dbWord," & Len(reqText) & ")='" & reqText & "';"
    'And Create a Recordset object with this SQLText
    Set RecSet = DBMain.OpenRecordset(SQLText)
    'Clear the list box
    Form1.lstWords.Clear
    If RecSet.RecordCount = 0 Then
        Form1.lstWords.AddItem "No item was found"
        Form1.txtExp.Text = ""
        Exit Sub
    End If
    'Do this, i will explain why later
    RecSet.MoveLast: RecSet.MoveFirst
    'fill the list box
    Do Until RecSet.EOF
        Form1.lstWords.AddItem RecSet.Fields(0)
        RecSet.MoveNext
    Loop
    'if there is just one word you see in the list
    'then pretend like user click a list box item
    'and bring the sigle item description which able to be selected
    If Form1.lstWords.ListCount = 1 Then
        RecSet.MoveFirst
        Form1.txtExp.Text = RecSet.Fields(1)
        Form1.txtWord.Text = Form1.lstWords.List(0)
        Form1.txtWord.SelLength = Len(Form1.txtWord.Text)
    Else
    'if not dont fill the list box
    'because there are lots of words
        Form1.txtExp.Text = ""
    End If

    'ok
    'This is the second code of mine
    'I see my code's visitor
    'But i am not really sure if you need my codes
    'If it is i will create bigger database application
    'here
    'Please tell me your idea.
    'Bye for today
End Sub

Public Sub CheckSameWord(reqText As String)
Dim SQLText As String
    'We are now checking the word which we press enter on text box
    'If we have 'Begin' and also 'Begining' explanations in database
    'and we also press Enter after we wrote 'Begin'
    'then query will check if that single word exists in database
    'If it is then it will be explained
    SQLText = "SELECT * FROM DBWords WHERE dbWord='" & reqText & "';" 'WHERE Left(dbWord," & Len(reqText) & ")='" & reqText & "';"
    Set RecSet = DBMain.OpenRecordset(SQLText)
    'If there is no record exit sub
    If RecSet.RecordCount = 0 Then Exit Sub
    Form1.txtExp.Text = RecSet.Fields(1)
    Form1.lstWords.Clear
    Form1.lstWords.AddItem RecSet.Fields(0)
    Form1.txtWord.SelStart = 0
    Form1.txtWord.SelLength = Len(Form1.txtWord.Text)
End Sub


Public Sub StoreData()
    'You've already got this procedure from the first code
End Sub
