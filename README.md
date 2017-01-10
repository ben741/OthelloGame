This is a simple implementation of Othello (also called Reversi) which you can play against an AI. 

## Installing and running

To play, run the file "Run_othello.exe". It should run fine on all versions of Windows. You may have to ignore some Windows security alerts.

## Playing

You always play as white, and go first. Just click where you want to place your piece. As soon as you place your piece, black will respond by playing their piece, and mark its move with a happy face.

## Rules

Your goal is to have more white pieces than black pieces at the end of the game. You do this by flipping the black pieces on your turn. Every black piece located on a continuous line (horizonal, vertical or diagonal) between the white piece you place and another white piece already on the board will be flipped. i.e. if you form a sequence WBBBW, with no gaps, it will be converted to WWWWWW. You are required to make a move that results in a flip, if such a move exists. A more detailed summary of the rules may be found at https://en.wikipedia.org/wiki/Reversi#Rules

I'm not sure if my program correctly handles the case where there are no available moves for white.

## How the AI works

As far as I can remember, the AI weights each square according to a predefined matrix. It knows that edges are good and corners are very good, but the squares next to the corners are bad (since they tend to give the opponent the corner). It uses those weights, and exploration of the game tree to a depth of 3 or so moves, to decide which move is best. In theory, it tries to either grab a good move right now, deny you a good move, or force you to open up a good move for it. At some point, I may do more archeology to figure out more details.
