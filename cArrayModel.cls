VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cArrayModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'################################################################################################################
'#                                              INSTRUCTIONS                                                    #
'################################################################################################################
'#                                                                                                              #
'#      In a standard module, dim the following variables:                                                      #
'#                                                                                                              #
'#  1. A storage array of undefined size (so it can be redimmed later)                                          #
'#                                                                                                              #
'#  2. A "New Entry" array of undefined size                                                                    #
'#                                                                                                              #
'#  3. An Array Model Object (for example, Dim AR as New cArrayModel)                                           #
'#                                                                                                              #
'#      Redim the storage array to have one row and columns equal to the number of properties.                  #
'#      Redim the New Entry array to be the same size as the columns in your storage array.                     #
'#                                                                                                              #
'#                                                                                                              #
'#  Example:                                                                                                    #
'#                                                                                                              #
'#  Dim StorAry() as Variant                                                                                    #
'#  Dim NewEntry() as Variant                                                                                   #
'#  Dim AR as New cArrayModel                                                                                   #
'#                                                                                                              #
'#  ReDim StorAry(1 to 1, 1 to 6)                                                                               #
'#  ReDim NewEntry(1 to 6)                                                                                      #
'#                                                                                                              #
'#                                                                                                              #
'#      Populate your Entry Array with information. Then add it to the array model.                             #
'#                                                                                                              #
'#  NewEntry(1) = "Sample"                                                                                      #
'#  NewEntry(2) = "Column Two"                                                                                  #
'#  NewEntry(3) = 8675309                                                                                       #
'#  NewEntry(4) = 42                                                                                            #
'#  NewEntry(5) = 4815162342                                                                                    #
'#  NewEntry(6) = TRUE                                                                                          #
'#                                                                                                              #
'#  Call AR.Add(StorAry, NewEntry)                                                                              #
'#                                                                                                              #
'#      The storage array now contains our record. Use a loop to print it out.                                  #
'#                                                                                                              #
'#  Dim I as Long                                                                                               #
'#  Dim Output as String                                                                                        #
'#                                                                                                              #
'#  For I = 1 to 6                                                                                              #
'#      Output = Output & StorAry(1, I) & vbCrLf                                                                #
'#  Next I                                                                                                      #
'#                                                                                                              #
'#  MsgBox Output                                                                                               #
'#                                                                                                              #
'#      You can continue adding records by populating NewEntry and using the Add method.                        #
'#  The array will automatically be resized as needed. There will probably be empty records                     #
'#  at the end of your array after adding. This is because the ReDim adds more than one                         #
'#  at a time to optimize run speed.                                                                            #
'#      After adding all your records, it's a good idea to run the .RemoveEmptyRows method.                     #
'#                                                                                                              #
'#  If you'd like to try a sample, copy the following code into a new standard Module and uncomment it.         #
'#  Make sure this entire file is entered as a class module named cArrayModel.                                  #
'#                                                                                                              #
'################################################################################################################

'Sub ArrayModelTest()
'
'    Dim StorAry() As Variant
'    Dim NewEntry() As Variant
'    Dim AR As New cArrayModel
'
'    ReDim StorAry(1 To 1, 1 To 6)
'    ReDim NewEntry(1 To 6)
'
'    NewEntry(1) = "Sample"
'    NewEntry(2) = "Column Two"
'    NewEntry(3) = 8675309
'    NewEntry(4) = 42
'    NewEntry(5) = 4815162342#
'    NewEntry(6) = True
'
'    Call AR.Add(StorAry, NewEntry)
'
'
'    Dim I As Long
'    Dim Output As String
'
'    For I = 1 To 6
'        Output = Output & StorAry(1, I) & vbCrLf
'    Next I
'
'    MsgBox Output
'
'End Sub



'################################################################################################################
'#                                              ARRAY MODEL                                                     #
'################################################################################################################
'#                                                                                                              #
'#      The Array Model is an alternative method for storing, sorting, and recalling records with multiple      #
'#  properties. This is functionally similar to both Collections and Dictionaries, but seeks to improve on      #
'#  their performance, which degrades sharply when reading and manipulating more than a few hundred records.    #
'#  The required syntax is different than these other objects due to VBA convention, but the Array Model        #
'#  retains the same functions as both of them while providing much faster processing.                          #
'#      You must format your data as an array of Variants, where the first index value represents the record    #
'#  index (same as Collecitions/Dictionaries), and the second index value represents the property of that       #
'#  record. You can enumerate the properties to make the code easier to read.                                   #
'#                                                                                                              #
'#      Note: Keys are not currently available. They may be added in a future version.                          #
'#                                                                                                              #
'################################################################################################################
'#                                                                                                              #
'#      AVAILABLE METHODS:                                                                                      #
'#                                                                                                              #
'#                                                                                                              #
'#      Add                                                                                                     #
'#      ---                                                                                                     #
'#                                                                                                              #
'#  Sub Add(ByRef Ary() As Variant, _                                                                           #
'#      ByRef Entry() As Variant, _                                                                             #
'#      Optional ByVal Before As Long = 0)                                                                      #
'#                                                                                                              #
'#      -Adds a new record to the array. If the size is not big enough, the method automatically ReDims it.     #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      DeepCopy                                                                                                #
'#      --------                                                                                                #
'#                                                                                                              #
'#  Sub DeepCopy(ByRef CopyAry() As Variant, _                                                                  #
'#               ByRef PasteAry() As Variant, _                                                                 #
'#               Optional ByVal StartRow As Long = 0, _                                                         #
'#               Optional ByVal EndRow As Long = 0, _                                                           #
'#               Optional ByVal RowOffset As Long = 0)                                                          #
'#                                                                                                              #
'#      -Copies the contents of CopyAry() to PasteAry(). By default it does the entire array, but it can be     #
'#       limited by passing a StartRow and EndRow. The RowOffset parameter allows you to move the contents      #
'#       up or down as they go to the paste array by a number of rows equal to the parameter.                   #
'#      (This is mostly meant for internal use, but may be useful in standard modules.)                         #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      Heapify                                                                                                 #
'#      -------                                                                                                 #
'#                                                                                                              #
'#  Sub Heapify(ByRef Ary() As Variant, Optional ByVal StartNode As Long = 0)                                   #
'#                                                                                                              #
'#      -Recursively applies the "Heap" Property to the array, such that record 1 is the Root node of a data    #
'#       tree and every parent node is greater than its children (MAX HEAP), or less than its children (MIN).   #
'#      -Note that this method should always be called after switching the property "HeapMode" true. Otherwise  #
'#       the existing records may not obey the Heap Property.                                                   #
'#      -This method was added specifically to make an A* pathfinding algorithm faster, as operating in Sorted  #
'#       mode takes too long to add nodes. Heaps are a standard for pathfinding.                                #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      QuickSort                                                                                               #
'#      ---------                                                                                               #
'#                                                                                                              #
'#  Sub QuickSort(ByRef Ary(), _                                                                                #
'#                ByVal StartIndex As Long, _                                                                   #
'#                ByVal EndIndex As Long)                                                                       #
'#                                                                                                              #
'#      -Sorts Ary using a very fast Quicksort algorithm. SortColumn is used as the sort key.                   #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      ReCount                                                                                                 #
'#      -------                                                                                                 #
'#                                                                                                              #
'#  Sub ReCount(ByRef Ary() As Variant)                                                                         #
'#                                                                                                              #
'#      -Checks the number of used and empty rows in the array and sets the Count and EmptyCount parameters     #
'#       based on what it finds. It's a good idea to call this method in the beginning, or any time you've      #
'#       changed the array size without using the Add or Remove methods.                                        #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      Remove                                                                                                  #
'#      ------                                                                                                  #
'#                                                                                                              #
'#  Sub Remove(ByRef Ary() As Variant, Index As Long)                                                           #
'#      -Removes the specified record from the array. Depending on the mode, it will either move the end        #
'#       record into the empty space or shift all the remaining records up one space. The first method of       #
'#       moving the end record into the empty is much faster but disrupts the sorting of the array.             #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      RemoveDuplicates                                                                                        #
'#      ----------------                                                                                        #
'#  Sub RemoveDuplicates(ByRef Ary() as Variant)                                                                #
'#      -Checks every entry in the array against every other entry. If a second entry is found with the         #
'#       exact same information in every field, delete the entry.                                               #
'#      -If the "KeepSorted" property is False, the sub has permission to sort the array using Quicksort to     #
'#       better optimize the search, but the array's order will probably change. If this property is True,      #
'#       instead run using a brute force approach. The Remove method will keep things sorted for us.            #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      RemoveEmptyRows                                                                                         #
'#      ---------------                                                                                         #
'#                                                                                                              #
'#  Sub RemoveEmptyRows(ByRef Ary() As Variant)                                                                 #
'#      -ReDims the array to remove any unused empty rows at the end. It's a good idea to call this before      #
'#       doing any output or pasting so that empty records aren't pasted.                                       #
'#                                                                                                              #
'#                                                                                                              #
'#                                                                                                              #
'#      SwapRecords                                                                                             #
'#      -----------                                                                                             #
'#                                                                                                              #
'#  Sub SwapRecords(ByRef Ary(), _                                                                              #
'#                  ByVal RecordOne As Long, _                                                                  #
'#                  ByVal RecordTwo As Long)                                                                    #
'#                                                                                                              #
'#      -Exchanges the position of two records in the array.                                                    #
'#      (This is mostly meant for internal use, but may be useful in standard modules.)                         #
'#                                                                                                              #
'################################################################################################################
'#                                                                                                              #
'#      AVAILABLE PROPERTIES:                                                                                   #
'#                                                                                                              #
'#                                                                                                              #
'#  Count (Long, Read-Only)                                                                                     #
'#      -This is the number of non-blank rows in the array. It keeps count when using the Add or Remove methods #
'#       or when you call the ReCount method.                                                                   #
'#                                                                                                              #
'#  EmptyCount (Long, Read-Only)                                                                                #
'#      -This is the number of blank rows at the end of the array. It keeps count in the same way as the Count  #
'#       property, and can be refreshed manually through the ReCount method.                                    #
'#                                                                                                              #
'#  HeapMode (Boolean, Read/Write, Default = FALSE)                                                             #
'#      -This optional switch forces the array model to operate in what is known as "Heap Mode". In this mode,  #
'#       records will be arranged to obey the Heap Property, in which the data is stored in a tree with the top #
'#       record (1) being either the Min or Max value of the data set, depending on if the Heap is a Min or     #
'#       Max. This is useful for always having access to the highest or lowest value, and was added to          #
'#       optimize the implementation of an A* Pathfinding algorithm.                                            #
'#      -You MUST call the Heapify method if you have any data in the array when you switch this on! Otherwise  #
'#       your data may not obey the Heap Property, and the implementation may fail.                             #
'#      -HeapMode sorts based on the SortColumn (default 1), which can be changed at any time. Be sure to call  #
'#       the Heapify method if you change the sort column.                                                      #
'#                                                                                                              #
'#  KeepSorted (Boolean, Read/Write, Default = FALSE)                                                           #
'#      -This optional switch forces the array model to insert new records based on the current sort order.     #
'#       It will run much slower in this mode, but the array will always be sorted on the SortColumn.           #
'#      -I don't suggest using this mode unless your algorithm depends on things always being sorted. Otherwise #
'#       it's faster to add all the records and then run QuickSort.                                             #
'#                                                                                                              #
'#  RedimSize (Long, Read/Write, Default = 0)                                                                   #
'#      -Determines how many blank spaces to add to the end of the storage array when trying to add a new       #
'#       record while there are no more blank spaces free. By default this is 1, but it can be increased to     #
'#       improve performance when adding a large number of records at once.                                     #
'#      -If setting RedimSize greater than 1, it's recommended to run RemoveEmptyRecords after adding new       #
'#       entries, to prevent having empty spaces in the array.                                                  #
'#      -If RedimSize is set to the default of 0, the array model will attempt to self-optimize by using an     #
'#       internal variable that increases proportional to the number of used records. Except in rare cases,     #
'#       I recommend leaving RedimSize at 0 and using the self-optimization.                                    #
'#      -REMEMBER TO RUN REMOVEEMPTYRECORDS WHEN DONE ADDING!                                                   #
'#                                                                                                              #
'#  SortAscending (Boolean, Read/Write, Default = TRUE)                                                         #
'#      -When the array is sorted using the QuickSort method, or if it's operating in KeepSorted mode, the      #
'#       lowest record will be in position 1. The SortColumn property is the sort key.                          #
'#                                                                                                              #
'#  SortDescending (Boolean, Read/Write, Default = FALSE)                                                       #
'#      -When the array is sorted using the QuickSort method, or if it's operating in KeepSorted mode, the      #
'#       highest record will be in position 1. The SortColumn property is the sort key.                         #
'#                                                                                                              #
'#  **** SortAscending and SortDescending are mutually excluslive. If you set one true, the other will    ****  #
'#  **** become False automatically, and vice-versa.                                                      ****  #
'#                                                                                                              #
'#  SortColumn (Long, Read/Write, Default = 1)                                                                  #
'#      -This is the column in the array (the second dimension position) that determines sort order. It is used #
'#       any time the QuickSort method is called, or if you're operating in KeepSorted mode.                    #
'#                                                                                                              #
'################################################################################################################
'#                              PERFORMANCE COMPARISON VS EXISTING OBJECTS                                      #
'################################################################################################################
'#                                                                                                              #                                                                                               #
'#  **Measurement Unit**                                                                                        #
'#  Writing a long value to a hard-coded variable has a performance of 1 LU (1 long unit time).                 #
'#                                                                                                              #
'#  **Array Model**                                                                                             #
'#  Writing a long value to an array of variants has a performance of 4 LU (takes four times as long).          #
'#                                                                                                              #
'#  **Traditional Dictionary or Collection**                                                                    #
'#  Writing a long value to a class object stored in a Colleciton or Dictionary has a performance of 60 LU      #
'#                                                                                                              #
'#      Using this metric, we see the write time is 60/4 = 15. Therefore we can write 15 entries to the Array   #
'#  Model in the time it takes to write one entry to a class object in a Collection or Dictionary.              #
'#                                                                                                              #
'#  >Actual testing shows performance closer to 5 times faster, not 15. Still significant.                      #
'#                                                                                                              #
'#  >The performance difference probably depends on the size of the second dimension in the array.              #
'#                                                                                                              #
'#  >Additionally, the performance suffers when operating in Sorted mode. The delay increases exponentially     #
'#   as the number of stored records increases. There is no simple way around this, but in cases where we need  #
'#   to keep the minimum or maximum record at the top of the array, a Heap algorithm will be added.             #
'#                                                                                                              #
'################################################################################################################


'################################################################################################################
'#                                              ERROR CODES                                                     #
'################################################################################################################
'#                                                                                                              #
'#  17001 - Calling procedure attempted to write to a read-only property.                                       #
'#                                                                                                              #
'#  17002 - Wrong number of parameters in an array passed to a procedure.                                       #
'#                                                                                                              #
'#  17003 - Index out of range.                                                                                 #
'#                                                                                                              #
'#  17004 - Value other than boolean was received by a sort parameter.                                          #
'#                                                                                                              #
'#  17005 - Calling procedure tried to sort a set of records greater than the maximum allowed. See Initialize.  #
'#                                                                                                              #
'#  17006 - Bad Index Number (Column) passed to SortAdvanced procedure. (Subscript out of range)                #
'#                                                                                                              #
'#  17007 - Invalid Sort Array passed to AdvancedSort procedure. (Array must be 2D, with exactly 2 fields)      #
'#                                                                                                              #
'################################################################################################################


'MEMBER VARIABLES'
'----------------'

Private m_lCount As Long
Private m_lEmptyCount As Long
Private m_bSortAscending As Boolean
Private m_bSortDescending As Boolean
Private m_lRedimSize As Long
Private m_lAutoRedimSize As Long
Private m_bKeepSorted As Boolean
Private m_bHeapMode As Boolean
Private m_bHeapMax As Boolean
Private m_bHeapMin As Boolean
Private m_lSortColumn As Long
Private m_bPreserveOrder As Boolean
'Internal Control Setting Variables - DEFAULT IS CONTROLLED WITHIN INITIALIZE METHOD
Private MaximumSortSize As Long
Private MaximumArraySize As Long


'PROPERTIES - GET/SET/LET STATEMENTS'
'-----------------------------------'

'The number of records stored in the array
Public Property Get Count() As Long
    Count = m_lCount
End Property
Public Property Let Count(lCount As Long)
    Err.Raise Number:=17001, Description:="Property 'Count' is Read-Only."
End Property

'The number of empty rows at the end of the array
Public Property Get EmptyCount() As Long
    EmptyCount = m_lEmptyCount
End Property
Public Property Let EmptyCount(lEmptyCount As Long)
    Err.Raise Number:=17001, Description:="Property 'EmptyCount' is Read-Only."
End Property

'Are the records sorted from low to high? (mutually exclusive with SortDescending)
Public Property Get SortAscending() As Boolean
    SortAscending = m_bSortAscending
End Property
Public Property Let SortAscending(bSortAscending As Boolean)
    'Ensure a proper boolean value is entered.
    If bSortAscending = True Or bSortAscending = False Then
        m_bSortAscending = bSortAscending
        If bSortAscending = True Then
            'Set SortDescending false because they can't both be true.
            m_bSortDescending = False
            m_bHeapMax = True
            m_bHeapMin = False
        Else
            'Set SortDescending true because the can't both be false.
            m_bSortDescending = True
            m_bHeapMax = False
            m_bHeapMin = True
        End If
    Else
        Err.Raise Number:=17004, Description:="Boolean value expected but not received for SortAscending property."
    End If
End Property

'Are the records sorted from high to low? (mutually exclusive with SortAscending)
Public Property Get SortDescending() As Boolean
    SortDescending = m_bSortDescending
End Property
Public Property Let SortDescending(bSortDescending As Boolean)
    'Ensure a proper boolean value is entered.
    If bSortDescending = True Or bSortDescending = False Then
        m_bSortDescending = bSortDescending
        If bSortDescending = True Then
            'Set SortAscending false because they can't both be true.
            m_bSortAscending = False
            m_bHeapMax = False
            m_bHeapMin = True
        Else
            'Set SortAscending true because they can't both be false.
            m_bSortAscending = True
            m_bHeapMax = True
            m_bHeapMin = False
        End If
    Else
        Err.Raise Number:=17004, Description:="Boolean value expected but not received for SortDescending property."
    End If
End Property

'Are we operating in Heap mode?
Public Property Get HeapMode() As Boolean
    HeapMode = m_bHeapMode
End Property
Public Property Let HeapMode(bHeapMode As Boolean)
    m_bHeapMode = bHeapMode
    'If HeapMode is being turned on, make sure "KeepSorted" is turned off.
    If bHeapMode = True Then
        m_bKeepSorted = False
    End If
End Property

'Set the heap property to obey MinHeap rules. (lowest value in root node)
Public Property Get HeapMin() As Boolean
    HeapMin = m_bHeapMin
End Property
Public Property Let HeapMin(bHeapMin As Boolean)
    m_bHeapMin = bHeapMin
    'Toggle HeapMax to the opposite.
    If bHeapMin = True Then
        m_bHeapMax = False
        m_bSortAscending = True
        m_bSortDescending = False
    Else
        m_bHeapMax = True
        m_bSortAscending = False
        m_bSortDescending = True
    End If
End Property

'Set the heap property to obey MaxHeap rules. (highest value in root node)
Public Property Get HeapMax() As Boolean
    HeapMax = m_bHeapMax
End Property
Public Property Let HeapMax(bHeapMax As Boolean)
    m_bHeapMax = bHeapMax
    'Toggle HeapMin to the opposite.
    If bHeapMax = True Then
        m_bHeapMin = False
        m_bSortAscending = False
        m_bSortDescending = True
    Else
        m_bHeapMin = True
        m_bSortAscending = True
        m_bSortAscending = False
    End If
End Property

'How much bigger do we make the array when we redim it?
Public Property Get RedimSize() As Long
    RedimSize = m_lRedimSize
End Property
Public Property Let RedimSize(lRedimSize As Long)
    m_lRedimSize = lRedimSize
End Property

'Should we keep the array sorted when adding & removing records?
Public Property Get KeepSorted() As Boolean
    KeepSorted = m_bKeepSorted
End Property
Public Property Let KeepSorted(bKeepSorted As Boolean)
    m_bKeepSorted = bKeepSorted
    'Have PreserveOrder follow KeepSorted. When it turns on or off, toggle
    'the other parameter with it.
    If bKeepSorted = True Then
        m_bPreserveOrder = True
        'Ensure Heap Mode is turned off.
        m_bHeapMode = False
    Else
        m_bPreserveOrder = False
    End If
End Property

'What is the column we want to sort the array on?
Public Property Get SortColumn() As Long
    SortColumn = m_lSortColumn
End Property
Public Property Let SortColumn(lSortColumn As Long)
    m_lSortColumn = lSortColumn
End Property

'Are we going to preserve the order of records in the array when removing one?
Public Property Get PreserveOrder() As Boolean
    PreserveOrder = m_bPreserveOrder
End Property
Public Property Let PreserveOrder(bPreserveOrder As Boolean)
    m_bPreserveOrder = bPreserveOrder
    'If PreserveOrder is switched off, also switch off KeepSorted.
    If bPreserveOrder = False Then m_bKeepSorted = False
End Property


'INITIALIZE'
'----------'

Private Sub Class_Initialize()
    'Set the sort option default to ascending mode.
    m_bSortAscending = True
    m_bSortDescending = False
    'Set the count of used and EmptyCount to zero.
    m_lCount = 0
    m_lEmptyCount = 0
    'Default Redim Size is 0, which means we'll use the auto redim feature.
    'User can disable by setting RedimSize greater than zero.
    m_lRedimSize = 0
    m_lAutoRedimSize = 1
    'Default sort column is 1
    m_lSortColumn = 1
    'Default sort mode is false.
    m_bKeepSorted = False
    'Default preserve mode is false.
    m_bPreserveOrder = False
    'Default Heap orientation is Max Heap.
    m_bHeapMax = False
    m_bHeapMin = True
    'Maxiumum sort size produces an error message telling the user they're
    'trying to sort too many rows.
    '*This is needed because Excel isn't allowed to allocate enough memory for
    '*such a large recursive function call. If the size goes above this, use
    '*the HeapSort algorthim instead. Almost as fast and never crashes.
    MaximumSortSize = 100
    'Maximum Array Size limits the expansion of the storage space in the Array Model.
    'This is used to prevent out of memory errors, but may result in a subscript out
    'of range error if the array goes over the maximum size.
    MaximumArraySize = 2000000
End Sub


'METHODS'
'-------'

'Load a new Entry into the first empty row of the Array.
Sub Add(ByRef Ary() As Variant, _
        ByRef Entry() As Variant, _
        Optional ByVal Before As Long = 0)
    'Temporary array in case we need to make the first one bigger.
    Dim TempAry() As Variant
    'First check if there are no records. In that case do an internal count.
    If m_lCount = 0 And m_lEmptyCount = 0 Then
        Call Me.ReCount(Ary)
    End If
    'If we're running in HeapMode, reset the Before value to 0. This is important
    'to maintain the Heap property, and we're going to add it to the end anyway.
    If m_bHeapMode = True Then Before = 0
    'If we're running in KeepSorted mode, search for the "Before" point and
    'overwrite anything already in that parameter.
    If m_bKeepSorted = True And m_lCount > 0 Then
        'The Binary Search was added as an inline function (No function call) to
        'save a small fraction of processing time.
        '**************************'
        '***INLINE BINARY SEARCH***'
        'Declare Two variables for doing the binary search.
        Dim TopBound As Long, BotBound As Long
        'Declare finishing variable
        Dim DoneSearching As Boolean
        DoneSearching = False
        'Check that our record isn't just going at the beginning or end.
        If m_bSortAscending = True Then
            'Our array is sorted ascending.
            If Entry(m_lSortColumn) <= Ary(1, m_lSortColumn) Then
                'This entry is less than the lowest value. Put at beginning.
                Before = 1
                DoneSearching = True
            Else
                If Entry(m_lSortColumn) >= Ary(m_lCount, m_lSortColumn) Then
                    'This entry is greater than the greatest value. Put at end.
                    Before = m_lCount + 1
                    DoneSearching = True
                End If
            End If
        Else
            'Our array is sorted descending.
            If Entry(m_lSortColumn) >= Ary(1, m_lSortColumn) Then
                'This entry is greater than the greatest value. Put at beginning.
                Before = 1
                DoneSearching = True
            Else
                If Entry(m_lSortColumn) <= Ary(m_lCount, m_lSortColumn) Then
                    'This entry is less than the lowest value. Put at end.
                    Before = m_lCount + 1
                    DoneSearching = True
                End If
            End If
        End If
        'If we haven't set the flag yet, that means we need to search inside.
        If DoneSearching = False Then
            'The entry falls somewhere in the middle. Find it.
            'Set the bottom and top bound and begin the search.
            TopBound = m_lCount
            BotBound = 1
            'Search for the record.
            Do
                Before = (TopBound + BotBound) \ 2
                'Check if we're seeing a match yet.
                Select Case Ary(Before, m_lSortColumn)
                    Case Is > Entry(m_lSortColumn)
                        'The new record is less than this place in the array.
                        If SortAscending = True Then
                            'We're sorted ascending. Ignore all entries less than this point.
                            BotBound = Before + 1
                        Else
                            'We're sorted descending. Ignore all entries greater than this point.
                            TopBound = Before - 1
                        End If
                    Case Is < Entry(m_lSortColumn)
                        'The new record is greater than this place in the array.
                        If SortAscending = True Then
                            'We're sorted ascending. Ignore all entries greater than this point.
                            TopBound = Before - 1
                        Else
                            'We're sorted descending. Ignore all entries less than this point.
                            BotBound = Before + 1
                        End If
                    Case Is = Entry(m_lSortColumn)
                        DoneSearching = True
                End Select
            Loop While DoneSearching = False And TopBound >= BotBound
            'We finished searching. Do a quick cleanup to make sure we're really pointed at the
            'record we want to insert before.
            If SortAscending = True Then
                'Sorted ascending. We should be pointed at a record that is greater than the entry,
                'and the one lower than that should be less than our entry.
                Do While Ary(Before, m_lSortColumn) < Entry(m_lSortColumn)
                    'This record is not greater than our new record. Loop until it is.
                    Before = Before + 1
                Loop
                Do While Ary(Before - 1, m_lSortColumn) > Entry(m_lSortColumn)
                    'The record one lower is not less than our new record. Loop until it is.
                    Before = Before - 1
                Loop
            Else
                'Sorted descending. We should be pointed at a record that is less than the entry,
                'and the one lower than that should be greater than our entry.
                Do While Ary(Before, m_lSortColumn) > Entry(m_lSortColumn)
                    'This record is not less than our new record. Loop until it is.
                    Before = Before + 1
                Loop
                Do While Ary(Before - 1, m_lSortColumn) < Entry(m_lSortColumn)
                    'The record one lower is not greater than our new record. Loop until it is.
                    Before = Before - 1
                Loop
            End If
        End If
        '***END OF BINARY SEARCH***'
        '**************************'
    End If
    'Make sure the insertion point is not greater than the number of used records plus one.
    If Before > m_lCount + 1 Then Before = m_lCount + 1
    'Check if the entry has the same number of columns as the array.
    If UBound(Ary, 2) = UBound(Entry) Then
        'Check if we need to add more empty rows.
        If m_lCount >= UBound(Ary, 1) Then
            'We need to resize the array to fit the new records. Load temp array.
            ReDim TempAry(UBound(Ary, 1), UBound(Ary, 2))
            'Deep copy to the temp array.
            Call Me.DeepCopy(Ary, TempAry)
            'Check if we're doing automatic redim size increases.
            If m_lRedimSize > 0 Then
                'Resize the array so it has space.
                ReDim Ary(1 To UBound(Ary, 1) + m_lRedimSize, 1 To UBound(Ary, 2))
                'Increase the count of unused rows.
                m_lEmptyCount = m_lRedimSize
            Else
                'Make sure the array is not being resized greater than the absolute limit of
                'the Array Model.
                If (UBound(Ary, 1) + m_lAutoRedimSize) * UBound(Ary, 2) > MaximumArraySize Then
                    'This redim will cause an overflow or out of memory crash. Avoid it by changing
                    'the Redim Size parameter.
                    m_lAutoRedimSize = CLng((MaximumArraySize / UBound(Ary, 2)) - (UBound(Ary, 1) + 1))
                End If
                'The redim size is inflating automatically. Use the other variable for this.
                ReDim Ary(1 To UBound(Ary, 1) + m_lAutoRedimSize, 1 To UBound(Ary, 2))
                'Increase the count of unused rows.
                m_lEmptyCount = m_lAutoRedimSize
                'Increase the redim size based on the number of records in the array.
                'Size increase is half of the count of used records plus one.
                m_lAutoRedimSize = CLng(m_lAutoRedimSize + (m_lCount / 2) + 1)
            End If
            'Deep copy back into the main array.
            Call Me.DeepCopy(TempAry, Ary)
        End If
        'Check if we're doing an insertion. If so, do a deep copy to create a space for
        'the record to go into.
        If Before > 0 Then
            'Move the records after the insert point all down one space, starting with
            'the end and working backward.
            Dim Index As Long
            Dim Col As Long
            For Index = m_lCount To Before Step -1
                For Col = 1 To UBound(Entry)
                    'Move this record one step down in the array.
                    Ary(Index + 1, Col) = Ary(Index, Col)
                Next Col
            Next Index
        Else
            'Since no insertion point was specified, use the first empty space in the array.
            Before = m_lCount + 1
        End If
        'Proceed with adding the new record to the array.
        Dim Param As Long
        For Param = 1 To UBound(Entry)
            Ary(Before, Param) = Entry(Param)
        Next Param
        'Increase the count of records used by one and reduce the count of
        'empty records by one.
        m_lCount = m_lCount + 1
        m_lEmptyCount = m_lEmptyCount - 1
        'Check if we're running in Heap Mode. If so, apply the heap property to this new record.
        If m_bHeapMode = True Then
            'Call the siftup method without the start node so it starts at the end of the array.
            Call Me.SiftUp(Ary)
        End If
    Else
        Err.Raise Number:=17002, Description:="Bad property count. Array and Entry must have the same number of properties (columns)." & _
                                              vbCrLf & "Array = " & UBound(Ary, 2) & ", Entry = " & UBound(Entry)
    End If
End Sub


'Remove an entry from the array.
Sub Remove(ByRef Ary() As Variant, Index As Long)
    'Indexing variable for the column, used in all branches of this procedure.
    Dim Param As Long
    'By default erase the record and move the one at the end into that empty space.
    'Start by validating the index number in case the calling function passed
    'a bad argument.
    If Index <= (LBound(Ary, 1) - 1) + m_lCount And Index >= LBound(Ary, 1) Then
        'Check if we're operating in Heap Mode or not.
        'By design the HeapMode flag overrides the PreserveOrder flag. It is not possible to be
        'both in PreserveOrder mode and obey the Heap Property (and you wouldn't want to).
        If m_bHeapMode = True Then
            'Since we're in Heap Mode we can skip most of the additional checks.
            'Swap this record with the very end of the array.
            Call Me.SwapRecords(Ary, Index, (LBound(Ary, 1) - 1) + m_lCount)
            'Remove the end record and decrease the order count by one.
            For Param = 1 To UBound(Ary, 2)
                Ary((LBound(Ary, 1) - 1) + m_lCount, Param) = Empty
            Next Param
            m_lCount = m_lCount - 1
            'Perform a Sift Down operation on the position we swapped, preserving the Heap Property.
            Call Me.SiftDown(Ary, Index)
        Else
            'Check if we're preserving order or not.
            If m_bPreserveOrder = True Then
                'Check if this record is the last one in the set.
                If Index = (LBound(Ary, 1) - 1) + m_lCount Then
                    'This is the end record. We can easily blank it out and be finished.
                    For Param = 1 To UBound(Ary, 2)
                        Ary(Index, Param) = Empty
                    Next Param
                Else
                    'Declare an empty storage array.
                    Dim TempAry() As Variant
                    'Make the new array the same size as the current one.
                    ReDim TempAry(UBound(Ary, 1), UBound(Ary, 2))
                    'Deep copy all the records before the index to the temp array.
                    Call Me.DeepCopy(Ary, TempAry, 1, Index - 1)
                    'Deep copy all the records after the index to the temp array.
                    'Offset by -1 to pull the rows together.
                    Call Me.DeepCopy(Ary, TempAry, Index + 1, (LBound(Ary, 1) - 1) + m_lCount, -1)
                    'Deep copy the temp array back into the original, overwriting it.
                    Call Me.DeepCopy(TempAry, Ary, LBound(Ary, 1), (LBound(Ary, 1) - 1) + m_lCount)
                End If
                'Reduce the used count by one
                m_lCount = m_lCount - 1
                'Increase the empty space by one
                m_lEmptyCount = m_lEmptyCount + 1
            Else
                'Since it is not necessary to preserve the order, knock out the selected record
                'and shuffle the last one into it.
                'Check if this is less than the last row.
                If Index < (LBound(Ary, 1) - 1) + m_lCount Then
                    'This is less than the last row. Move the last record into the empty space.
                    For Param = 1 To UBound(Ary, 2)
                        Ary(Index, Param) = Ary((LBound(Ary, 1) - 1) + m_lCount, Param)
                    Next Param
                End If
                'Blank out the last row
                For Param = 1 To UBound(Ary, 2)
                    Ary((LBound(Ary, 1) - 1) + m_lCount, Param) = Empty
                Next Param
                'Reduce the used count by one
                m_lCount = m_lCount - 1
                'Increase the empty space by one
                m_lEmptyCount = m_lEmptyCount + 1
            End If
        End If
    Else
        Err.Raise Number:=17003, Description:="Index out of range. Index Value = " & Index & ", " & LBound(Ary) & " Max = " & (LBound(Ary) - 1) + m_lCount
    End If
End Sub

'Checks every row of the array in the model and removes any rows with exactly the same data.
Sub RemoveDuplicates(Ary() As Variant)
    'Use the array model internally to handle records going into a new array.
    Dim CarMod As New cArrayModel
    Dim RecordsToRemove() As Variant
    Dim NewRecordToRemove() As Variant
    ReDim RecordsToRemove(1 To 1, 1 To 1)
    ReDim NewRecordToRemove(1 To 1)
    Call CarMod.ReCount(RecordsToRemove)
    'Declare local variables.
    Dim CurrentRecord As Long 'The record we're using as the unique entry.
    Dim Index As Long 'The record we're checking to see if it's a duplicate.
    Dim ExactMatch As Boolean 'Tells us if the Index location is an exact match for all fields or not.
    Dim Col As Long 'Iterates through every Field (column) of the array.
    'Start by checking if we're in KeepSorted mode or not.
    If m_bPreserveOrder = True Then
        'The calling function cares that the order does not change. Use a brute force approach.
        CurrentRecord = 1
        Do While CurrentRecord < m_lCount
            Index = 1
            'Check the Current Record aginst every Index location in the array.
            Do While Index <= m_lCount
                'Check if the Current Record matches every Field (column) of the Index location.
                'Only do this if our current record and index are different locations.
                If CurrentRecord <> Index Then
                    ExactMatch = True
                    'Check each Field. If any is not a match, set ExactMatch to false and break early.
                    For Col = 1 To UBound(Ary, 2)
                        If Ary(CurrentRecord, Col) <> Ary(Index, Col) Then
                            'This is not a matching record. break early and set the flag False.
                            ExactMatch = False
                            Exit For
                        End If
                    Next Col
                    'Did we get an exact match?
                    If ExactMatch = True Then
                        'Since this is an exact match, we don't care about indexing the record. We'll remove this
                        'duplicate and m_lCount will be reduced by one.
                        Call Me.Remove(Ary, Index)
                    Else
                        'Index the current record being examined to continue checking.
                        Index = Index + 1
                    End If
                Else
                    Index = Index + 1
                End If
            Loop
            'Go to the next record in order.
            CurrentRecord = CurrentRecord + 1
        Loop
    Else
        'The calling function doesn't care about the order of the data. Sort ascending so we can step
        'down more efficiently.
        Call Me.QuickSort(Ary)
        CurrentRecord = 1
        Index = 2
        Do While CurrentRecord < m_lCount And Index <= m_lCount
            'Enter a loop while the Index record has the same sort key value as the CurrentRecord.
            Do While Ary(CurrentRecord, m_lSortColumn) = Ary(Index, m_lSortColumn)
                'Skip this record if the current record and index are the same.
                If CurrentRecord <> Index Then
                    'If all the fields match between these records, put them on the remove list.
                    ExactMatch = True
                    For Col = 1 To UBound(Ary, 2)
                        If Ary(CurrentRecord, Col) <> Ary(Index, Col) Then
                            'This is not a matching record. break early and set the flag False.
                            ExactMatch = False
                            Exit For
                        End If
                    Next Col
                    'Check if this should go on the list.
                    'FORGET THE LIST! HE GOES TO THE BLOCK! Remember that game? Email CWillard.net@gmail.com subject: "Fus Ro Dah"
                    'I'd love to know that you're using my code and paying attention to the comments.
                    If ExactMatch = True Then
                        'Add this to the array model for the list of rows we'll be deleting after we're done.
                        NewRecordToRemove(1) = Index
                        Call CarMod.Add(RecordsToRemove, NewRecordToRemove)
                    End If
                End If
                'Index to the next record.
                Index = Index + 1
                'If we've gone over the count of records, exit the loop early.
                If Index > m_lCount Then Exit Do
            Loop
            'Step down to the next record that isn't on the duplicates list.
            ExactMatch = True
            Do While ExactMatch = True
                'Step down to the next entry.
                CurrentRecord = CurrentRecord + 1
                'Make sure we're still in the array somewhere.
                If CurrentRecord > m_lCount Then Exit Do
                'Loop to check each entry in the "delete these" array. If we find a match to our current
                'row, break early. If we found no match, set the ExactMatch flag false and let us exit.
                ExactMatch = False
                For Index = 1 To CarMod.Count
                    If RecordsToRemove(Index, 1) = CurrentRecord Then
                        'This record is already set to be deleted. Skip it by setting ExactMatch back to "True"
                        ExactMatch = True
                        Exit For
                    End If
                Next Index
            Loop
            'Move the Index value to the next row after our current record.
            Index = CurrentRecord + 1
        Loop
        'Remove all the records called out on the list, from the LAST to the FIRST.
        'We need to do it this way because every removal causes the last entry in the array to replace the
        'entry being removed. If we went from top to bottom there's a risk that the end of the array could get
        'unsorted, and a duplicate would be put into the middle.
        'By working from the end we can be certain of removing all the records before they're mis-sorted.
        For Index = CarMod.Count To 1 Step -1
            CurrentRecord = RecordsToRemove(Index, 1)
            Call Me.Remove(Ary, CurrentRecord)
            Debug.Print "Debug"
        Next Index
    End If
End Sub

'Copies the contents of one array into another.
Sub DeepCopy(ByRef CopyAry() As Variant, _
             ByRef PasteAry() As Variant, _
             Optional ByVal StartRow As Long = 0, _
             Optional ByVal EndRow As Long = 0, _
             Optional ByVal RowOffset As Long = 0)
    Dim Row As Long
    Dim Col As Long
    'Set values for the start and end rows if needed.
    If StartRow = 0 Then StartRow = 1
    If EndRow = 0 Then EndRow = UBound(CopyAry, 1)
    If UBound(PasteAry, 1) >= EndRow And _
       UBound(PasteAry, 2) >= UBound(CopyAry, 2) Then
        For Row = StartRow To EndRow
            For Col = 1 To UBound(CopyAry, 2)
                'Use the appropriate copy mode depending on variable type.
                If VarType(CopyAry(Row, Col)) <> vbObject Then
                    PasteAry(Row + RowOffset, Col) = CopyAry(Row, Col)
                Else
                    'This is an object, so we need to use the "Set" keyword.
                    Set PasteAry(Row + RowOffset, Col) = CopyAry(Row, Col)
                End If
            Next Col
        Next Row
    Else
        Err.Raise Number:=17002, Description:="Bad property count. PasteAry must be the same size" & _
                                              " or larger than the number of records being copied."
    End If
End Sub


'Quckly and recursively sort the array.
Sub QuickSort(ByRef Ary(), _
              Optional ByVal StartIndex As Long = 0, _
              Optional ByVal EndIndex As Long = 0)
    'Quicksort procedure selects the rightmost value as a pivot, and puts anything
    'less than that on the left, and everything greater than that on the right.
    'It runs recursively, cutting the array in half each time, until the array size
    'is less than two.
    '***********************************************************************************************'
    '*** This is based off a c++ algorithm found at the following site: ***'
    'https://medium.com/human-in-a-machine-world/quicksort-the-best-sorting-algorithm-6ab461b5a9d0
    '*** The syntax was rewritten to follow VBA code standards. ***'
    '***********************************************************************************************'
    Dim Cnt As Long
    Dim ArryStep As Long
    Dim Pivot As Variant
    'If the end index was not specified, set it to the maximum used area of the array in the model.
    If StartIndex = 0 Then StartIndex = 1
    If EndIndex = 0 Then EndIndex = m_lCount
    'Make sure the count of records in the array is not greater than the maximum allowed.
    If m_lCount > MaximumSortSize Then
        'This array is too large for recursive sorting. We will have to use a slightly slower HeapSort
        'algorithm instead. It will Heapify and then Sort the array.
        '*NOTE: AFTER TESTING I FOUND THIS TO BE JUST AS FAST AS THE QUICKSORT*'
        Me.HeapMode = False
        Call Me.HeapSort(Ary, StartIndex, EndIndex)
        Exit Sub
    End If
    'Operate the sort based on the sort direction.
    If m_bSortAscending = True Then
        'Do the sort from low to high.
        If StartIndex < EndIndex Then
            '"Cnt" indicates the split point between elements smaller than the pivot
            'and elements greater than the pivot.
            Cnt = StartIndex
            'Set the pivot value equal to the value of the EndIndex item.
            Pivot = Ary(EndIndex, m_lSortColumn)
            'Go through the entire array
            For ArryStep = StartIndex To EndIndex
                'Check the element to see if it's less than the pivot value
                If Ary(ArryStep, m_lSortColumn) <= Pivot Then
                    'This value is less than the pivot. Place it on the left side of the pivot value.
                    'Do this if the records are not in the same position!
                    Call Me.SwapRecords(Ary(), ArryStep, Cnt)
                    'Increment the Cnt value so we'll be able to swap the next small value
                    'with the next "untouched" value.
                    Cnt = Cnt + 1
                End If
            Next ArryStep
            'At this point our Cnt value is at the position immediately following our
            'pivot value. We'll need to back it up by 2 places for the next iterative loop.
            '*Recursively sort left side of pivot*
            Call Me.QuickSort(Ary(), StartIndex, Cnt - 2)
            '*Recursively sort right side of pivot*
            Call Me.QuickSort(Ary(), Cnt, EndIndex)
        End If
    Else
        'Do the sort from high to low.
        If StartIndex < EndIndex Then
            '"Cnt" indicates the split point between elements smaller than the pivot
            'and elements greater than the pivot.
            Cnt = StartIndex
            'Set the pivot value equal to the value of the EndIndex item.
            Pivot = Ary(EndIndex, m_lSortColumn)
            'Go through the entire array
            For ArryStep = StartIndex To EndIndex
                'Check the element to see if it's greater than the pivot value
                If Ary(ArryStep, m_lSortColumn) >= Pivot Then
                    'This value is greater than the pivot. Place it on the left side of the pivot value.
                    Call Me.SwapRecords(Ary(), ArryStep, Cnt)
                    'Increment the Cnt value so we'll be able to swap the next big value
                    'with the next "untouched" value.
                    Cnt = Cnt + 1
                End If
            Next ArryStep
            'At this point our Cnt value is at the position immediately following our
            'pivot value. We'll need to back it up by 2 places for the next iterative loop.
            '*Recursively sort left side of pivot*
            Call Me.QuickSort(Ary(), StartIndex, Cnt - 2)
            '*Recursively sort right side of pivot*
            Call Me.QuickSort(Ary(), Cnt, EndIndex)
        End If
    End If
End Sub


'Exchange the position of two records in the array.
Sub SwapRecords(ByRef Ary(), _
                ByVal RecordOne As Long, _
                ByVal RecordTwo As Long)
    'This function takes an array and the position of two records.
    'It swaps these records by using a second internal array.
    Dim StoreAry() As Variant
    Dim Index As Long
    Dim KeyOne As String
    Dim KeyTwo As String
    ReDim StoreAry(1 To UBound(Ary, 2))
    'Only do the swap if the records are not in the same position.
    If RecordOne <> RecordTwo Then
        'Swap record one into the storage array, then record two into
        'record one, and finally the storage array into record two.
        For Index = 1 To UBound(Ary, 2)
            'Swap Record One into the storage array.
            If VarType(Ary(RecordOne, Index)) <> vbObject Then
                StoreAry(Index) = Ary(RecordOne, Index)
            Else
                Set StoreAry(Index) = Ary(RecordOne, Index)
            End If
            'Swap Record Two into Record One
            If VarType(Ary(RecordTwo, Index)) <> vbObject Then
                Ary(RecordOne, Index) = Ary(RecordTwo, Index)
            Else
                Set Ary(RecordOne, Index) = Ary(RecordTwo, Index)
            End If
            'Swap Storage Array into Record One
            If VarType(StoreAry(Index)) <> vbObject Then
                Ary(RecordTwo, Index) = StoreAry(Index)
            Else
                Set Ary(RecordTwo, Index) = StoreAry(Index)
            End If
        Next Index
    End If
End Sub


'Remove any empty spaces at the end of the array.
Sub RemoveEmptyRows(ByRef Ary() As Variant)
    'Make a temporary storage array.
    Dim TempAry() As Variant
    'Set up the array.
    ReDim TempAry(LBound(Ary, 1) To (LBound(Ary, 1) - 1) + m_lCount, 1 To UBound(Ary, 2))
    'Deep copy only the non-blank records to the storage array.
    Call Me.DeepCopy(Ary, TempAry, LBound(Ary, 1), (LBound(Ary, 1) - 1) + m_lCount)
    'Blank out the original array by redimming it.
    ReDim Ary(LBound(Ary, 1) To (LBound(Ary, 1) - 1) + m_lCount, 1 To UBound(Ary, 2))
    'Recopy the records back over.
    Call Me.DeepCopy(TempAry, Ary, LBound(Ary, 1), (LBound(Ary, 1) - 1) + m_lCount)
    'Set the internal empties variable to zero.
    m_lEmptyCount = 0
End Sub


'Recalculates the Count and EmptyCount Properties.
Sub ReCount(ByRef Ary() As Variant)
    'Iterate through the array and count how many empty and non-empty rows there are.
    Dim Row As Long, Col As Long
    Dim EmptyRow As Boolean
    'Blank out the array model counts.
    m_lCount = 0
    m_lEmptyCount = 0
    For Row = LBound(Ary, 1) To UBound(Ary, 1)
        EmptyRow = True
        For Col = LBound(Ary, 2) To UBound(Ary, 2)
            If Ary(Row, Col) <> Empty Then
                EmptyRow = False
                Exit For
            End If
        Next Col
        If EmptyRow = True Then
            'When you hit the first empty, do the count calculation.
            Exit For
        End If
    Next Row
    'Row is now pointing at the first empty row. If it's beyond the maximum size of the array,
    'this means there are no empties.
    If Row > UBound(Ary, 1) Then
        'All rows are full.
        m_lCount = UBound(Ary, 1) - (LBound(Ary, 1) - 1)
    Else
        'Used rows equal the first empty row minus one.
        m_lCount = Row - 1 - (LBound(Ary, 1) - 1)
        'Empty rows equal the size of the array minus used rows.
        m_lEmptyCount = (UBound(Ary, 1) - (LBound(Ary, 1) - 1)) - m_lCount
    End If
End Sub


'Takes an array and performs a Heapify operation on it.
'Note: THIS IS NOT A HEAP SORT! This will take an existing array and do sift-up, sift-down
'      operations until the array satisfies the Heap property.
'YOU MUST CALL THIS METHOD IF YOU ADDED RECORDS BEFORE TURNING ON HEAP MODE!
Sub Heapify(ByRef Ary() As Variant, _
            Optional ByVal RootNode As Long = 0, _
            Optional ByVal EndNode As Long = 0, _
            Optional ByVal OffsetNode As Long = 1)
    'Iterate through all used records in the array and sort them recursively based on the Heap Property.
    'Check if RootNode has children, and call Heapify for both of them.
    '********************************************************************************************************'
    '*  The optional SubHeap Flag allows the function to treat the Start to End nodes as if they're an      *'
    '*  independent array. "Start Node" will be treated as the root in this case, and the formula           *'
    '*  that identifies child nodes will be modified to change the structure of the heap.                   *'
    '********************************************************************************************************'
    Dim SwapNode As Long 'The SwapNode is the one we're going to exchange for the root.
    Dim LeftChild As Long
    Dim RightChild As Long
    Dim LeftOffsetNode As Long
    Dim RightOffsetNode As Long
    Dim SubRootNode As Long
    'If Heapify is called without a start node, set it to 1.
    '*THIS IS THE DEFAULT BEHAVIOR, SO THE ENTIRE ARRAY IS SORTED BASED ON THE ROOT NODE!*'
    If RootNode = 0 Then RootNode = 1
    '*THE DEFAULT END NODE IS THE COUNT OF RECORDS IN THE ARRAY MODEL*'
    If EndNode = 0 Then EndNode = m_lCount
    '*LEFT CHILD CAN BE FOUND BY ADDING THE ROOT NODE WITH THE OFFSET*'
    LeftChild = RootNode + OffsetNode
    '*LEFT CHILD WILL HAVE AN OFFSET NODE EQUAL TO DOUBLE THE CURRENT OFFSET*'
    LeftOffsetNode = OffsetNode * 2
    '*RIGHT CHILD IS ALWAYS ONE HIGHER THAN THE LEFT CHILD*'
    RightChild = LeftChild + 1
    '*RIGHT CHILD WILL HAVE AN OFFSET NODE EQUAL TO THE LEFT OFFSET NODE PLUS ONE*'
    RightOffsetNode = LeftOffsetNode + 1
    'RECURSIVELY SIFT THROUGH BOTH DIRECTIONS'
    '*RECURSIVELY CALL LEFT CHILD*'
    If LeftChild <= EndNode Then
        Call Me.Heapify(Ary, LeftChild, EndNode, LeftOffsetNode)
    End If
    '*RECURSIVELY CALL RIGHT CHILD*'
    If RightChild <= EndNode Then
        Call Me.Heapify(Ary, RightChild, EndNode, RightOffsetNode)
    End If
    'This node is now ready to be heaped. Branch based on our type of heaping.
    If m_bHeapMax = True Then
        '* MAX HEAP *' - The highest value must go to the Root Node.
        'Branch based on how far down in the tree we are.
        Select Case EndNode
            Case Is >= RightChild
                'Both child nodes are occupied.
                'Choose the higher of the two nodes to be our swapper.
                If Ary(LeftChild, m_lSortColumn) > Ary(RightChild, m_lSortColumn) Then
                    SwapNode = LeftChild
                    SubRootNode = LeftOffsetNode
                Else
                    SwapNode = RightChild
                    SubRootNode = RightOffsetNode
                End If
                'Check if our swap node is greater than the parent node.
                If Ary(SwapNode, m_lSortColumn) > Ary(RootNode, m_lSortColumn) Then
                    Call Me.SwapRecords(Ary, SwapNode, RootNode)
                    'Since we swapped records, recursively call Heapify for the branch we swapped.
                    Call Me.Heapify(Ary, SwapNode, EndNode, SubRootNode)
                End If
            Case Is = LeftChild
                'Only left child node is occupied.
                SwapNode = LeftChild
                SubRootNode = LeftOffsetNode
                'Check if our swap node is greater than the parent node.
                If Ary(SwapNode, m_lSortColumn) > Ary(RootNode, m_lSortColumn) Then
                    Call Me.SwapRecords(Ary, SwapNode, RootNode)
                    'Since we swapped records, recursively call Heapify for the branch we swapped.
                    Call Me.Heapify(Ary, SwapNode, EndNode, SubRootNode)
                End If
            Case Else
                'Neither node is occupied. Do nothing.
        End Select
    Else
        '* MIN HEAP *' - The lowest value must go to the Root Node.
        'Branch based on how far down in the tree we are.
        Select Case EndNode
            Case Is >= RightChild
                'Both child nodes are occupied.
                'Choose the lower of the two nodes to be our swapper.
                If Ary(LeftChild, m_lSortColumn) < Ary(RightChild, m_lSortColumn) Then
                    SwapNode = LeftChild
                    SubRootNode = LeftOffsetNode
                Else
                    SwapNode = RightChild
                    SubRootNode = RightOffsetNode
                End If
                'Check if our swap node is less than the parent node.
                If Ary(SwapNode, m_lSortColumn) < Ary(RootNode, m_lSortColumn) Then
                    Call Me.SwapRecords(Ary, SwapNode, RootNode)
                    'Since we swapped records, recursively call Heapify for the branch we swapped.
                    Call Me.Heapify(Ary, SwapNode, EndNode, SubRootNode)
                End If
            Case Is = LeftChild
                'Only left child node is occupied.
                SwapNode = LeftChild
                SubRootNode = LeftOffsetNode
                'Check if our swap node is less than the parent node.
                If Ary(SwapNode, m_lSortColumn) < Ary(RootNode, m_lSortColumn) Then
                    Call Me.SwapRecords(Ary, SwapNode, RootNode)
                    'Since we swapped records, recursively call Heapify for the branch we swapped.
                    Call Me.Heapify(Ary, SwapNode, EndNode, SubRootNode)
                End If
            Case Else
                'Neither node is occupied. Do nothing.
        End Select
    End If
End Sub


'Takes a Heapified array and sorts it.
'NOTE: The array's sort order will be opposite of the Heap Type. For example, if you start with
'      a Max Heap, your result will be sorted in Ascending (low to high) order. If you start with
'      a Min Heap, your result will be sorted in Descending (high to low) order.
Sub HeapSort(ByRef Ary() As Variant, _
             Optional ByVal StartIndex As Long = 0, _
             Optional ByVal EndIndex As Long = 0)
    'The Heap Sort allows a non-recursive method for sorting the array model. It is not as
    'fast as the QuickSort method, but can be used for arrays exceeding the maximum size allowed
    'by QuickSort (Usually around 20,000 entries, depending on Excel version and system memory.)
    'This Sub is called automatically if the count of entries exceeds the max QuickSort size.
    Dim HeapSize As Long
    'Check if we're using default values for start and end.
    If StartIndex = 0 Then StartIndex = 1
    If EndIndex = 0 Then EndIndex = m_lCount
    HeapSize = EndIndex 'Start our count for the main loop based on the number of records we have.
    'Check first to see if we're operating in Heap Mode. If not we'll have to Heapify first.
    If m_bHeapMode = False Then Call Me.Heapify(Ary, StartIndex, EndIndex, 1)
    'Run a loop until our HeapSize has shrunk enough to show us the array is fully sorted.
    Do While HeapSize > StartIndex
        'Exchange the root node with the bottom of the array.
        Call Me.SwapRecords(Ary, StartIndex, HeapSize)
        'Shrink HeapSize by 1.
        HeapSize = HeapSize - 1
        'Re-Heapify the array through a quick Siftdown if we have more than one record in the heap.
        If HeapSize > StartIndex Then
            Call Me.SiftDown(Ary, StartIndex, HeapSize)
        End If
    Loop
    'The array is now sorted. Turn off Heap
End Sub


'Takes a position in a heapified array and examines the nested parent nodes recursively
'until the node and its parent all obey the heap property.
Sub SiftUp(ByRef Ary() As Variant, _
           Optional StartNode As Long = 0, _
           Optional EndNode As Long = 0, _
           Optional OffsetNode As Long = 1)
    Dim ParentNode As Long
    'If the StartNode was omitted, set it equal to the maximum of the array.
    If StartNode = 0 Then StartNode = m_lCount
    'If the StartNode is greater than 1 (the root node), perform the sift check.
    If StartNode > 1 Then
        'Compare the Start Node with its parent node.
        If StartNode Mod 2 = 0 Then
            'The starting node is an even value. Its Parent is half its value.
            ParentNode = StartNode / 2
        Else
            'The starting node is an odd value. Its Parent is half its value minus one.
            ParentNode = (StartNode - 1) / 2
        End If
        'Branch based on our type of heap.
        If HeapMin = True Then
            'This is a Min heap. Check if our node is less than its parent.
            If Ary(StartNode, m_lSortColumn) < Ary(ParentNode, m_lSortColumn) Then
                'This node does not obey the Heap property. Exchange the values.
                Call Me.SwapRecords(Ary, StartNode, ParentNode)
                'Since we had to swap, it's possible the parent now violates the Heap Property.
                'Recursively call this function to check that step using the parent node.
                Call Me.SiftUp(Ary, ParentNode)
            End If
        Else
            'This is a Max heap. Check if our node is greater than its parent.
            If Ary(StartNode, m_lSortColumn) > Ary(ParentNode, m_lSortColumn) Then
                'This node does not obey the Heap property. Exchange the values.
                Call Me.SwapRecords(Ary, StartNode, ParentNode)
                'Since we had to swap, it's possible the parent now violates the Heap Property.
                'Recursively call this function to check that step using the parent node.
                Call Me.SiftUp(Ary, ParentNode)
            End If
        End If
    End If
End Sub


'Takes a position in a heapified array and examines the child nodes recursively
'until the node and its children all obey the heap property.
Sub SiftDown(ByRef Ary() As Variant, _
             Optional StartNode As Long = 0, _
             Optional EndNode As Long = 0, _
             Optional OffsetNode As Long = 1)
    'Behavioral note:
    'The algorithm will prefer the right branch if both left and right are equal and the
    'Heap Property is not satisfied. This should not affect general use, but is important
    'to understand that the sorting will generally be slightly more accurate than if
    'the left branch was preferred.
    Dim ChildNode As Long
    Dim BranchMove As Long
    Dim LeftChild As Long
    Dim RightChild As Long
    Dim NextOffsetNodeLeft As Long
    Dim NextOffsetNodeRight As Long
    'Set up the Left and Right Child positions and prep the next offset node.
    LeftChild = StartNode + OffsetNode
    RightChild = StartNode + OffsetNode + 1
    NextOffsetNodeLeft = OffsetNode * 2
    NextOffsetNodeRight = NextOffsetNodeLeft + 1
    'Our Branchmove tells us which way we're going. It is zero if the parent and children
    'satisfy the Heap property, 1 if we need to swap and go LEFT, or 2 if we need to
    'swap and go RIGHT.
    BranchMove = 0
    'If the StartNode was omitted, set it equal to the root position of the array.
    If StartNode = 0 Then StartNode = 1
    'If the EndNode was omitted, set it equal to the count of records in the array.
    If EndNode = 0 Then EndNode = m_lCount
    'If the StartNode is less than or equals the count (the final child), perform the sift check.
    If StartNode <= EndNode Then
        'Branch based on our heap type.
        If m_bHeapMin = True Then
            'MIN HEAP. Look for the lowest value among the children and parent.
            'First check to make sure both right and left children are populated nodes.
            If RightChild <= EndNode Then
                'Check if parent is greater than left child.
                If Ary(StartNode, m_lSortColumn) > Ary(LeftChild, m_lSortColumn) Then
                    'The parent is greater than the left child. increase BranchMove by one.
                    BranchMove = 1
                    'Check if the left child is greater than the right child. If so we swap them.
                    If Ary(LeftChild, m_lSortColumn) >= Ary(RightChild, m_lSortColumn) Then
                        'our right child is the greatest (or equal and we prefer the right branch).
                        'We will change our anticipated direction from left to right.
                        BranchMove = 2
                    End If
                Else
                    'Since the parent is not greater than left child, check the right child.
                    If Ary(StartNode, m_lSortColumn) > Ary(RightChild, m_lSortColumn) Then
                        'The right branch is greater than parent node. Prepare to swap these.
                        BranchMove = 2
                    End If
                End If
            Else
                'Since the right node is not populated, check if the left node is populated.
                'Note: if both right and left are not populated this means we've reached the bottom
                'of our sifting and will pass a BranchMove of 0 to the next check.
                If LeftChild <= EndNode Then
                    'There's something in the left node. Check if parent node is greater.
                    If Ary(StartNode, m_lSortColumn) > Ary(LeftChild, m_lSortColumn) Then
                        'Parent node is greater. Prepare to swap them.
                        BranchMove = 1
                    End If
                End If
            End If
        Else
            'MAX HEAP. Look for the lowest value between node and children.
            'First check to make sure both right and left children are populated nodes.
            If RightChild <= EndNode Then
                'Check if parent is less than left child.
                If Ary(StartNode, m_lSortColumn) < Ary(LeftChild, m_lSortColumn) Then
                    'The parent is greater than the left child. increase BranchMove by one.
                    BranchMove = 1
                    'Check if the left child is greater than the right child.
                    If Ary(LeftChild, m_lSortColumn) <= Ary(RightChild, m_lSortColumn) Then
                        'our right child is the greatest (or equal and we prefer the right branch).
                        'We will change our anticipated direction from left to right.
                        BranchMove = 2
                    End If
                Else
                    'Since the parent is not less than left child, check the right child.
                    If Ary(StartNode, m_lSortColumn) < Ary(RightChild, m_lSortColumn) Then
                        'The right branch is greater than parent node. Prepare to swap these.
                        BranchMove = 2
                    End If
                End If
            Else
                'Since the right node is not populated, check if the left node is populated.
                'Note: if both right and left are not populated this means we've reached the bottom
                'of our sifting and will pass a BranchMove of 0 to the next check.
                If LeftChild <= EndNode Then
                    'There's something in the left node. Check if parent node is less.
                    If Ary(StartNode, m_lSortColumn) < Ary(LeftChild, m_lSortColumn) Then
                        'Parent node is greater. Prepare to swap them.
                        BranchMove = 1
                    End If
                End If
            End If
        End If
        'Now that a direction is chosen, branch to the code that matches.
        Select Case BranchMove
            Case Is = 1
                'We must swap parent with the LEFT node and recursively call that direction.
                Call Me.SwapRecords(Ary, StartNode, LeftChild)
                'Call the procedure again.
                Call Me.SiftDown(Ary, LeftChild, EndNode, NextOffsetNodeLeft)
            Case Is = 2
                'We must swap parent with the RIGHT node and recursively call that direction.
                Call Me.SwapRecords(Ary, StartNode, RightChild)
                'Cal the procedure again.
                Call Me.SiftDown(Ary, RightChild, EndNode, NextOffsetNodeRight)
        End Select
    End If
End Sub

Sub SortAdvanced(ByRef Ary() As Variant, _
                 ByRef SortColumnAndDirection() As Variant, _
                 Optional ByVal StartIndex As Long = 0, _
                 Optional ByVal EndIndex As Long = 0)
    'Sorts the array with an arbitrary number of nested sorts. "Sort By this, then that, then that, etc..."
    Dim SortLevel As Long
    Dim StartSort As Long
    Dim EndSort As Long
    Dim Index As Long
    'Set the defaults for start and end indexes.
    If StartIndex = 0 Then StartIndex = LBound(Ary, 1)
    If EndIndex = 0 Then EndIndex = UBound(Ary, 1)
    'Validate the sort array - it must be 2d with exactly 2 fields (SortColumn, SortDirection) where SortDirection = TRUE
    'means sorting Ascending and SortDirection = FALSE means sorting Descending.
    On Error Resume Next
    If UBound(SortColumnAndDirection, 2) <> 2 Then
        On Error GoTo 0
        Err.Raise Number:=17007, _
                  Description:="Error: Invalid SortColumnAndDirection() array received by SortAdvanced procedure in Array Model. " & _
                               "Array must have exactly two columns." & vbCrLf & vbCrLf & "(N, 1) = Sort Columns." & _
                               vbCrLf & "(N, 2) = Sort Directions where TRUE = Ascending."
    End If
    'Start by doing the first Sort Level.
    SortLevel = LBound(SortColumnAndDirection, 1)
    If SortColumnAndDirection(SortLevel, 1) >= LBound(Ary, 2) And SortColumnAndDirection(SortLevel, 1) <= UBound(Ary, 2) Then
        m_lSortColumn = SortColumnAndDirection(SortLevel, 1)
        Select Case SortColumnAndDirection(SortLevel, 2)
            Case True
                Me.SortAscending = True
            Case False
                Me.SortAscending = False
        End Select
    Else
        Err.Raise Number:=17006, _
                  Description:="Subscript above or below maximum allowed for array when calling SortAdvanced function." & _
                               " Received: " & SortColumnAndDirection(SortLevel, 1) & ", Min: " & LBound(Ary, 2) & _
                               ", Max: " & UBound(Ary, 2) & ", Step: " & SortLevel
    End If
    'Subscript is good. Call the sort for this first level.
    Call Me.QuickSort(Ary, StartIndex, EndIndex)
    'Begin looping through all remaining subscripts in the ordered sort columns array.
    For SortLevel = LBound(SortColumnAndDirection, 1) + 1 To UBound(SortColumnAndDirection, 1)
        'Check the subscript to make sure it's an allowed value.
        If SortColumnAndDirection(SortLevel, 1) >= LBound(Ary, 2) And SortColumnAndDirection(SortLevel, 1) <= UBound(Ary, 2) Then
            m_lSortColumn = SortColumnAndDirection(SortLevel, 1)
            Select Case SortColumnAndDirection(SortLevel, 2)
                Case True
                    Me.SortAscending = True
                Case False
                    Me.SortAscending = False
            End Select
        Else
            Err.Raise Number:=17006, _
                      Description:="Subscript above or below maximum allowed for array when calling SortAdvanced function." & _
                                   " Received: " & SortColumnAndDirection(SortLevel, 1) & ", Min: " & LBound(Ary, 2) & _
                                   ", Max: " & UBound(Ary, 2) & ", Step: " & SortLevel
        End If
        'Perform a series of sorts by stepping through the previous level sort and creating a group.
        Dim Entry As Long
        'Always start the first subset at the top.
        StartSort = StartIndex
        For Entry = StartIndex To EndIndex
            'Check if the Entry is already at the end index.
            If Entry = EndIndex Then
                'Subsort the last section of this sort level.
                EndSort = Entry
                'Check to make sure we're sorting at least 2 entries.
                If EndSort - StartSort > 0 Then
                    Call Me.QuickSort(Ary, StartSort, EndSort)
                End If
            Else
                'Check each row in order until you find one that doesn't match.
                'We must check all previous subscripts until we find one that isn't an exact match. Otherwise it's possible
                'to incorrectly sort where a record looked identical but was actually different.
                Dim ExactMatch As Boolean
                ExactMatch = True
                'Step through every previously sorted column until you find one where the previously sorted column
                'is not the same as a previous record.
                For Index = 1 To SortLevel - 1
                    If Ary(Entry, SortColumnAndDirection(Index, 1)) <> Ary(Entry + 1, SortColumnAndDirection(Index, 1)) Then ExactMatch = False
                Next Index
                'We're checking the previous subscript in the ordered set because it determines the limits of our next sort.
                If ExactMatch = False Then
                    'We've found the first sorted entry that doesn't match the previous row. Do a subsort here.
                    EndSort = Entry
                    'Check to make sure we're sorting at least 2 entries.
                    If EndSort - StartSort > 0 Then
                        Call Me.QuickSort(Ary, StartSort, EndSort)
                    End If
                    'Step the start sort down further.
                    StartSort = Entry + 1
                End If
            End If
        Next Entry
    Next SortLevel
End Sub

Sub ClearUsedRange(ByVal WS As Worksheet, _
                   Optional ByVal HasHeaders As Boolean = False)
    'Simple support procedure, takes a worksheet object and clears the entire used range
    'through the implicit Worksheet.Usedrange() property.
    '(NOTE: THERE ARE SOME CASES OF THIS PROPERTY NOT BEING COMPLETELY ACCURATE! IN THESE
    ' CASES THE CLEAR SHOULD STILL WORK FINE, ONLY DELETING EMPTY CELLS.)
    Dim FirstRow As Long
    Dim LastRow As Long
    Dim FirstCol As Long
    Dim LastCol As Long
    'Select the second row if the worksheet contains headers.
    If HasHeaders = True Then
        FirstRow = 2
    Else
        FirstRow = 1
    End If
    'Always start on column one.
    FirstCol = 1
    'Manage any errors that resulted from an invalid worksheet.
    On Error GoTo ErrorHandler
    LastRow = WS.UsedRange.Rows.Count
    LastCol = WS.UsedRange.Columns.Count
    If LastRow >= FirstRow Then
        WS.Range(WS.Cells(FirstRow, FirstCol), _
                 WS.Cells(LastRow, LastCol)).Delete xlShiftUp
    End If
    'Exit without error.
    Exit Sub
ErrorHandler:
    Err.Raise Number:=Err.Number, _
              Description:="Encountered within ClearUsedRange procedure. Error code native to VBA." & vbCrLf & vbCrLf & _
              "Error Text: " & Err.Description
End Sub

Sub PasteArrayValuesToWorksheet(ByRef Ary() As Variant, _
                                ByRef WS As Worksheet, _
                                Optional ByVal StartRow As Long = 1, _
                                Optional ByVal StartCol As Long = 1)
    'Dumps the array into a worksheet, starting at the specified position.
    WS.Range(WS.Cells(StartRow, StartCol), _
             WS.Cells(UBound(Ary, 1) + StartRow - 1, UBound(Ary, 2) + StartCol - 1)).Value = Ary
End Sub

Function FindBin(ByVal Term As Variant, _
                 ByRef Ary() As Variant, _
                 ByVal SearchColumn As Long, _
                 Optional ByVal StartRow As Long = 0, _
                 Optional ByVal EndRow As Long = 0) As Long
    'Performs a binary search and returns the index number of the record
    'that matches the Term provided in the specified SearchColumn of the supplied array, Ary().
    Dim TBound As Long
    Dim BBound As Long
    Dim FoundMatch As Boolean
    'Set the top and bottom of the search range. Uses the supplied optional variables, or
    'the top and bottom of the array if they're omitted.
    If StartRow > 0 Then
        BBound = StartRow
    Else
        BBound = LBound(Ary, 1)
    End If
    If EndRow > 0 Then
        TBound = EndRow
    Else
        TBound = UBound(Ary, 1)
    End If
    'Set initial condition.
    FoundMatch = False
    'Begin the search.
    Do
        'Set the search target to the middle of the remaining range.
        FindBin = (TBound + BBound) \ 2
        'Check if we have a match in this position.
        Select Case Ary(FindBin, SearchColumn)
            Case Is > Term
                'This entry is greater than our search term. Limit the upper bound.
                TBound = FindBin - 1
            Case Is < Term
                'This entry is greater than our search term. Limit the lower bound.
                BBound = FindBin + 1
            Case Is = Term
                FoundMatch = True
        End Select
    Loop While TBound >= BBound And FoundMatch = False
    'If we failed to find a matching term, return zero.
    If FoundMatch = False Then FindBin = 0
End Function









