Sub AllCaps()  
    For Each Cell In Selection  
        If Not Cell.HasFormula Then  
            Cell.Value = UCase(Cell.Value)  
        End If  
    Next Cell  
End Sub  

' __       __    __     .___  ___.  __   __  ___  __  .______           ___________    __    ____    .___  ___.  __   __  ___  __  .______      
'|  |     |  |  |  |    |   \/   | |  | |  |/  / |  | |   _  \         /  _____\   \  /  \  /   /    |   \/   | |  | |  |/  / |  | |   _  \     
'|  |     |  |  |  |    |  \  /  | |  | |  '  /  |  | |  |_)  |       |  |  __  \   \/    \/   /     |  \  /  | |  | |  '  /  |  | |  |_)  |    
'|  |     |  |  |  |    |  |\/|  | |  | |    <   |  | |      /        |  | |_ |  \            /      |  |\/|  | |  | |    <   |  | |      /     
'|  `----.|  `--'  |    |  |  |  | |  | |  .  \  |  | |  |\  \----.   |  |__| |   \    /\    /       |  |  |  | |  | |  .  \  |  | |  |\  \----.
'|_______| \______/     |__|  |__| |__| |__|\__\ |__| | _| `._____|    \______|    \__/  \__/        |__|  |__| |__| |__|\__\ |__| | _| `._____|
                                                                                                                                               '
