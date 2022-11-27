# Experience_Viewer
Tool to edit "eman-like" EXP file

Prerequisites :<br>
copy [pgn-extract.exe](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/pgn-extract.exe) to your EXPERIENCE VIEWER folder<p>

rename BUREAU.ini to YOUR-COMPUTER-NAME.ini<br>
set moteurEXP to your_eman_engine.exe<br>
set fichierEXP to your_experience_file.exp<p>
  
# How it works ?
There are different ways to use this tool :<br>
- enter an UCI string or left click on one of the moves present in the experience data<br>
![uci_string](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/uci_string.jpg)<p>

- enter an EPD string<br>
![epd_string](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/epd_string.jpg)<p>

- click on the "..." button to load a PGN or an EPD file<p>

# How to edit experience data ?
- If you want to delete a move from the experience data, double-click on its move's index<p>

- If you want to change the move's depth, double-click on its move score/depth<p>

- If you want to change the move's count, double-click on its move count<p>

- If you want to add more moves to the current position or update the evaluation of the learned moves, click on the "Analyse" button.<br>
Then, configure your engine and click on the "Apply" button<br>
![configure_engine](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/configure_engine.jpg)<br>
Then, choose your analysis mode and click on the "ok" button<br>
![choose_mode](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/choose_mode.jpg)<p>

# tips
Right-click on the large gray text box to copy its contents to the clipboard.<p>
![clipboard](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/clipboard.jpg)<p>

After you load a PGN, use the "<<" and ">>" buttons to change the game and the "<" and ">" buttons to change the move.<br>
![pgn_file](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/pgn_file.jpg)<p>

After you load an EPD file, use the "<" and ">" buttons to change the position.<br>
![epd_file](https://github.com/chris13300/experience_viewer/blob/main/Experience%20Viewer/bin/x64/Debug/epd_file.jpg)<p>

Which analysis modes to use ?<br>
all legal moves : it corresponds to the same needs as a MultiPV analysis.<br>
only played moves : it is useful for improving the evaluation of the learned moves, aligning their depths, etc.<br>
only moves with positive scores : it is useful in positions containing a lot of learned moves so as not to reanalyze the worst moves.<br>
search the bestmove : it corresponds to the same needs as a fixed-depth analysis.<p>
