<div align="center">

## Tabstrip project


</div>

### Description

This is a tabstrip project. By using an array, you can decide which tab container should be shown. Whatever you want to show when that tab is clicked goes into a container. I searched for months trying to learn about tabstrips and I hope this code helps someone else out.
 
### More Info
 
In order to use the tabstrip control you must make

each tab a seperate container. For four tabs, you

need four containers. By using a control array, you

scroll through the containers depending on which

tab you have selected. For this project I used picture

boxes as the containers. In the form load procedure, the

containers are made hidden so that only the one selected

is visible. Also, the border is set to zero at form load.

This is because when you're working on the containers,

it's easier if you can see the border. At run time,

you don't need the container to show, only the items

you put into the container.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David VanHook](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-vanhook.md)
**Level**          |Unknown
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-vanhook-tabstrip-project__1-2406/archive/master.zip)





### Source Code

```
For this project you will need:
1 Form - People
1 Command button - cmdexit
1 TabStrip - TabStrip1 (default)
 Place 4 tabs onto the tabstrip
4 Pictureboxes (in an array)
 A) Picture1(1)
 B) Picture1(2)
 C) Picture1(3)
 D) Picture1(4)
Const Numtabs = 4 'Set the number of tabs
Dim x as Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdexit_Click()
 Unload People
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
 On Error Resume Next
 People.Height = 3375 'Set the size of your form
 People.Width = 4900
 For x = 1 To Numtabs 'Loop through the tabs
 With Picture1(x)
 .BorderStyle = 0
 .Left = TabStrip1.ClientLeft
 .Top = TabStrip1.ClientTop
 .Width = TabStrip1.ClientWidth
 .Height = TabStrip1.ClientHeight
 .Visible = False
 End With
 Next x
 TabStrip1.Tabs(1).Selected = True 'Form loads with first tab selected
 Picture1(TabStrip1.SelectedItem.Index).Visible = True 'Show first container
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TabStrip1_Click()
 'This procedure determines which tab is selected
 'and what tab container should be shown
 Static PrevTab As Integer
 PrevTab = Switch(PrevTab = 0, 1, PrevTab >= 1 And PrevTab <= Numtabs, PrevTab)
 Picture1(PrevTab).Visible = False
 Picture1(TabStrip1.SelectedItem.Index).Visible = True
 Picture1(TabStrip1.SelectedItem.Index).Refresh
 PrevTab = TabStrip1.SelectedItem.Index
End Sub
'If you have any questions or problems, contact me:
'Zombiehead@earthlink.net
'http://home.earthlink.net/~zombiehead/vbexamples.htm
```

