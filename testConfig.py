
import ConfigParser

config = ConfigParser.ConfigParser()
config.read('config.conf')
prefreq = config.get('config','preFreq')
if '' == prefreq :
    print 1
else :
    print 2