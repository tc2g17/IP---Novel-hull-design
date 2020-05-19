import win32com.client
import numpy as np
import matplotlib.pyplot as plt
import time

"""
VARIABLES
"""
seastates = []#characteristic wave height, wave period probability of occurance

VCG = 0 #set in meters from baseline
speed =0.51444* 30 #in m/s
headings = np.linspace(0, np.pi, 13) # radians
colours = ['k','b', 'g', 'r', 'y', 'm', 'c']

"""
LIMITS
"""
time = 360 #estimated journey time in minutes
sickness = 5 # % of sickness incidence onboard
k = 1/3
def v_a_limit(time, k, sickness): #time in minutes, k as constat typically taken as 1/3
    acc = sickness/(k*time)
    return acc

vertical_acceleration_limit = v_a_limit(time, k, sickness)
horizontal_acceleration_limit = 9.81*0.025
roll_limit = (np.pi/180)*3
pitch_limit = (np.pi/180)*2

"""
ANALYSIS
"""
 
msApp = win32com.client.Dispatch("BentleyModeler.Application")
skApp = win32com.client.Dispatch("BentleyMotions.Application")
path1 = r""
path2 = r""
path3 = r""
path4 = r""
path5 = r""
hulls = [path1, path2, path3, path4, path5]


def go(hull):
    path=hulls[hull-1]
    start_time = time.time()
    msApp.Design.Open(path, False, False) #working. opens ms design
    msApp.Design.Hydrostatics.Calculate()
    skApp.Design.DesignOpen(path)
    Length = msApp.Design.Hydrostatics.LWL
    Beam = msApp.Design.Hydrostatics.BeamWL
    
    skApp.Design.Vessel.GyradiusPitch = Length*0.25 #from variables section
    skApp.Design.Vessel.GyradiusRoll = Beam*0.4
    skApp.Design.Vessel.GyradiusYaw = Length*0.25
    skApp.Design.Vessel.NumOfMappingTerms = 3
    skApp.Design.Vessel.NumOfSections = 41
    skApp.Design.Vessel.RollDampingNonDim = 0.1
    skApp.Design.Vessel.vcg = VCG
    skApp.Design.CalculateGeometry()
    
    skApp.Design.AnalysisOptions.NumOfFrequencies = 100
    skApp.Design.AnalysisOptions.UseTransomTerms = False
    skApp.Design.AnalysisOptions.WaterDensity = 1025
    
    skApp.Design.Speeds.RemoveAll()
    skApp.Design.Headings.RemoveAll()
    skApp.Design.Spectra.RemoveAll()
    
    skApp.Design.Speeds.Add("Std speed", speed) #speed in m/s
    
    for i in headings:
        skApp.Design.Headings.Add("heading {}".format(i), i) #heading in radians

    
    for i in range(len(seastates)):
        skApp.Design.Spectra.Add()
        skApp.Design.Spectra.Item(i+1).Analyse = True
        skApp.Design.Spectra.Item(i+1).name = "sea {}".format(i+1)
        skApp.Design.Spectra.Item(i+1).CharacteristicWaveHeight = seastates[i][0]
        skApp.Design.Spectra.Item(i+1).ModalPeriod = seastates[i][1]
        
    skApp.Design.CalculateSeakeeping()
    
    """
    RESULTS
    """
    #speed/heading/spectrum
    rms_vertical_acceleration=[]
    roll=[]
    pitch = []
    for j in range(len(headings)):
        AccRMS = []
        rollRMS = []
        pitchRMS = []
        for i in range(len(seastates)):
            indi_Acc = skApp.Design.GlobalStatistics.Item(1, j+1, i+1).HeaveAcceleration_rms
            indi_roll = skApp.Design.GlobalStatistics.Item(1, j+1, i+1).RollMotion_rms
            indi_pitch = skApp.Design.GlobalStatistics.Item(1, j+1, i+1).PitchMotion_rms
            AccRMS.append(indi_Acc)
            rollRMS.append(indi_roll)
            pitchRMS.append(indi_pitch)
        rms_vertical_acceleration.append(AccRMS)
        roll.append(rollRMS)
        pitch.append(pitchRMS)

    
    print("--- %s seconds ---" % (time.time() - start_time))
    return rms_vertical_acceleration, roll, pitch

def rms_v_limits(rms_vertical_acceleration):
    limiting_seastate=[]
    for j in rms_vertical_acceleration:
        above=[]
        for i in range(len(j)):
            if j[i]>vertical_acceleration_limit:
                above.append((j[i], i))
        limiting_seastate.append(above)

    if len(seastates[0])==3:
        vertical_acceleration_opreability=[]
        for j in limiting_seastate:
            probability=[]
            for i in j:
                probability.append(seastates[i[1]][2])
            vertical_acceleration_opreability.append(100-sum(probability))
    else:
        vertical_acceleration_opreability=[]
        for i in limiting_seastate:
            probability=(1-len(i)/len(seastates))*100
            vertical_acceleration_opreability.append(probability)
    print('VERT', vertical_acceleration_opreability)
    return vertical_acceleration_opreability

def roll_limits(roll):
    limiting_seastate=[]
    for j in roll:
        above=[]
        for i in range(len(j)):
            if j[i]>roll_limit:
                above.append((j[i], i))
        limiting_seastate.append(above)

    if len(seastates[0])==3:
        roll_operability=[]
        for j in limiting_seastate:
            probability=[]
            for i in j:
                probability.append(seastates[i[1]][2])
            roll_operability.append(100-sum(probability))
    else:    
        roll_operability=[]
        for i in limiting_seastate:
            probability=(1-len(i)/len(seastates))*100
            roll_operability.append(probability)
    print('ROLL', roll_operability)
    return roll_operability

def pitch_limits(pitch):
    limiting_seastate=[]
    for j in pitch:
        above=[]
        for i in range(len(j)):
            if j[i]>pitch_limit:
                above.append((j[i], i))
        limiting_seastate.append(above)

    if len(seastates[0])==3:
        pitch_operability=[]
        for j in limiting_seastate:
            probability=[]
            for i in j:
                probability.append(seastates[i[1]][2])
            pitch_operability.append(100-sum(probability))
    else:
        pitch_operability=[]
        for i in limiting_seastate:
            probability=(1-len(i)/len(seastates))*100
            pitch_operability.append(probability)
    print('PITCH', pitch_operability)
    return pitch_operability
            
def operability(rms_vertical_acceleration, roll, pitch):
    operability_test1 = [min(value) for value in zip(roll_limits(roll), rms_v_limits(rms_vertical_acceleration))]
    operability = [min(value) for value in zip(operability_test1, pitch_limits(pitch))]

    return operability

def results(hull):
    rms_vertical_acceleration, roll, pitch = go(hull)
    results = [rms_vertical_acceleration, roll, pitch]
    return results

def plot(hull):
    rms_vertical_acceleration, roll, pitch = results(hull);
    op = operability(rms_vertical_acceleration, roll, pitch)
    headings2=[i+np.pi for i in headings]
    op2=op[::-1]
    ax = plt.subplot(111, projection='polar')
    ax.plot(headings, op, color='b')
    ax.plot(headings2, op2, color='b')
    ax.set_xticks(np.linspace(0, 2*np.pi, 25))
    ax.set_rmax(100)
    ax.grid(True)
    ax.set_theta_zero_location('S')
    ax.set_title("Operability of test vessel as % in average annual wave data", va='bottom')
    plt.show()
    
def plot_all():
    ax = plt.subplot(111, projection='polar')
    for i in range(len(hulls)):
        rms_vertical_acceleration, roll, pitch = results(i+1);
        op = operability(rms_vertical_acceleration, roll, pitch)
        headings2=[i+np.pi for i in headings]
        op2=op[::-1]
        ax.plot(headings, op, color= colours[i], label='Hull{}'.format(i+1))
        ax.plot(headings2, op2, color=colours[i])
    ax.legend()
    ax.grid(True)
    ax.set_xticks(np.linspace(0, 2*np.pi, 25))
    ax.set_rmax(100)
    ax.set_theta_zero_location('S')
    #ax.set_title("Operability of test vessel as % in average annual wave data", va='bottom')
    plt.show()
