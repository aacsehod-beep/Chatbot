"""Inject logo base64 into index_new.html → templates/index.html"""
import re

src  = open('templates/index.html', encoding='utf-8').read()
imgs = re.findall(r'data:image/png;base64,[A-Za-z0-9+/=]+', src)
LOGO    = imgs[0]
LOGO_SM = imgs[1] if len(imgs) > 1 else imgs[0]

tpl = open('templates/index_new.html', encoding='utf-8').read()
out = tpl.replace('__LOGO__', LOGO).replace('__LOGO_SM__', LOGO_SM)

open('templates/index.html', 'w', encoding='utf-8').write(out)
f = open('templates/index.html', encoding='utf-8').read()
print('Lines:', f.count('\n')+1)
print('Lucide:', 'lucide' in f)
print('base64:', 'base64' in f)
print('All JS:', all(x in f for x in ['sendMessage','toggleDark','selectService',
      'clearChat','toggleVoice','renderSuggestions','useSuggestion',
      'guessModule','scrollToBottom','startChat']))
import os; os.remove('templates/index_new.html')
print('Done.')
