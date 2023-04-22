cd ~/Medvind

while true; do
   echo "Run Medvind"
   ~/Medvind/venv/bin/python run.py
   echo "   ... Medvind done"
   perl -le 'sleep rand 7200'
done

