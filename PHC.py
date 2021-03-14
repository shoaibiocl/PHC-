import salabim as sim
import random
import numpy as np
import xlsxwriter
import matplotlib.pyplot as plt
import statistics
import xlwt

""" PHC operational model at minutes time scale. From the class Main OPD arrivals and admin work arrival are generated.
OPD runs for 6 hours in a day followed by 2 hours of admin work for doctors.
"""


class Main(sim.Component):
    global OPD_iat
    global fail_count
    global ncd_time


    env = sim.Environment()

    ncd_time=0
    No_of_shifts = 0                                    # tracks number of shifts completed during the simulation time
    No_of_days = 0
    NCD_admin_work = []                                     # array for storing the admin work generated
    SN_time = 0                                         # for recording staff nurse time
    staff_nurse_del = 0
    staff_nurse_ANC = 0
    staff_nurse_IPD = 0
    p_count =0
    NT_list= []      # to store nurse time
    fail_count = 0

    warm_up = 180*24*60
    def process(self):

        global ncd_time

        self.shift = 0
        self.sim_time = 0                               # local variable defined for dividing each day into shits
        self.z=0
        self.admin_count = 0
        k = 0

        while self.z % 3 == 0:                          # condition to run simulation for 8 hour shifts
            Main.No_of_days += 1                        # class variable to track number of days passed
            while self.sim_time < 360:                  # from morning 8 am to afternoon 2 pm (in minutes)
                if env.now() <= Main.warm_up:
                    pass
                else:
                    Main.p_count +=1
                OPD_PatientGenerator()
                patient_ia_time = sim.Exponential(OPD_iat).sample()  # Patient inter arrival time
                yield self.hold(patient_ia_time)
                self.sim_time += patient_ia_time

            while 360 <= self.sim_time < 480:            # condition for admin work after opd hours are over
                k = int(sim.Normal(100,20).bounded_sample(60,140))
                """For sensitivity analysis. Admin work added to staff nurse rather than doctor"""
                if env.now() <= Main.warm_up:
                    pass
                else:
                    Patient.doc_service_time.append(k) # conatns all doctor service times
                    Main.NCD_admin_work.append(k)
                    ncd_time += k
                yield self.hold(120)
                self.sim_time = 481
            self.z += 3
            self.sim_time=0
            Main.No_of_shifts += 3
            yield self.hold(960)                        # holds simulation for 2 shifts


class OPD_PatientGenerator(sim.Component):

    """This is called from main. It assigns attributes to the patients based on which their flow is determined."""
    OPD_List = {}                                       # log of all the patient stored here
    patient_count = 0                                   # total OPD patients
    patient_count_1 = 0                                 # total OPD patients with single visit
    patient_count_2 = 0                                 # total OPD patients with single visit
    patient_count_3 = 0                                 # total OPD patients with single visit

    def __init__(self):                                 # calls process method every time OPD patient class is called
        self.dic = {}                                   # local dictionary for temporarily   storing generated patients
        # with attributes
        OPD_PatientGenerator.patient_count += 1
        self.time_of_visit = [[0], [0], [0]]            # initializing for assigning time of visits
        self.registration_time = round(env.now())       # registration time is the current simulation time
        self.id = OPD_PatientGenerator.patient_count    # patient count is patient id
        self.age_random = random.randint(0, 1000)       # assigning age to population based on census 2011 gurgaon
        if self.age_random <= 578:
            self.age = random.randint(0, 30)
        else:
            self.age = random.randint(31, 100)
        self.sex = random.choice(["male", "female"])                # Equal probability of male and female patient
        self.type_of_patient = random.choice(["True", "False"])     # Considering nearly half of all the patients
        # require lab tests
        self.visits_assigned = random.randint(1, 10)                # assigning number of visits visits
        self.dic = {"ID":self.id, "Age": self.age, "Sex": self.sex, "Lab": self.type_of_patient,
                    "Registration_Time": self.registration_time, "No_of_visits": self.visits_assigned,
                    "Time_of_visit": self.time_of_visit}
        OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count] = self.dic
        self.process()

    def process(self):

        x = OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["No_of_visits"]
        """x checks for the number of visits a patient will make."""

        if x == 2 or x == 3:                                    # 20 % patients assigned 2 visits
            OPD_PatientGenerator.patient_count_2 += 1
            y = random.randint(3*24*60, 8*24*60)                # patient can visit between 3 and 8 days
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][0] = \
                round(self.registration_time)
            # calling patient class for 1st visit
            Patient()
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][1] \
                = round(self.registration_time + y)
            # scheduling patient class for 2nd visit
            Patient(at=OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][1])
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][2] = float("inf")

        elif x == 1:                                            # 10% patients assigned 3 visits
            OPD_PatientGenerator.patient_count_3 += 1
            y = random.randint(3*24*60, 8*24*60)                # patient can visit between 3 and 8 days for 2nd visit
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][0] = \
                round(self.registration_time)
            Patient()
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][1] = \
                round(self.registration_time + y)
            Patient(at=OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][1])
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][2] = \
                round(self.registration_time + 2 * y)           # patient can visit between 3 and 8 days after 2nd visit
            Patient(at=OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][2])

        else:
            OPD_PatientGenerator.patient_count_1 += 1           # patients with only 1 visit
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][0] = \
                round(self.registration_time)
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][1] = float("inf")
            OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Time_of_visit"][2] = float("inf")
            Patient()


"""IPD patients are generated using a separate generator. Since IPD patients are emergency cases, hence they can
come at time during a day.
"""


class IPD_PatientGenerator(sim.Component):

    global IPD_iat
    IPD_List = {}                                       # log of all the IPD patients stored here
    patient_count = 0
    p_count = 0                                         # log of patients in each replication

    def process(self):
        global M
        M = 0
        while True:
            if env.now() <= Main.warm_up:
                pass
            else:
                IPD_PatientGenerator.patient_count += 1
                IPD_PatientGenerator.p_count += 1
            self.registration_time = env.now()
            self.id = IPD_PatientGenerator.patient_count
            self.age = round(random.normalvariate(35,8))
            self.sex = random.choice(["Male", "Female"])
            IPD_PatientGenerator.IPD_List[self.id] = [self.registration_time, self.id, self.age, self.sex]
            if 0 < (self.registration_time - N * 1440) < 480:
                IPD_with_doc(urgent=True)
            else:
                IPD_no_doc(urgent=True)
            self.hold_time_1 = sim.Exponential(IPD_iat).sample()
            yield self.hold(self.hold_time_1)
            M = int(env.now()/1440)


"""This class models the delivery patients in a PHC. The patient arrival is divided into two parts: 1) During OPD hours
  and 2) during non OPD hours. It is divided into OPD and non OPD hours. The delivery patient arrival can happen
  anytime during a day. A delivery patient gets priority over others and they are treated as urgent. If arrival occurs
   during the OPD hours then the Delivery_with_doc class is instantiated otherwise the other one without the doctor"""


class Delivery(sim.Component):

    global delivery_iat
    global N

    Delivery_list = {}
    Delivery_count = 0
    p_count = 0     # to maintain the log of patients in each replication

    N = 0

    def process(self):
        global N

        while True:
            if env.now() <= Main.warm_up:
                pass
            else:
                Delivery.p_count += 1
                Delivery.Delivery_count += 1
            self.registration_time = round(env.now())
            self.id = Delivery.Delivery_count
            self.sex = "Female"
            Delivery.Delivery_list[self.id] = [self.registration_time, self.id, self.sex]
            if 0 < (self.registration_time - N * 1440) < 480:
                Delivery_with_doctor(urgent=True)      # sets priority
            else:
                Delivery_no_doc(urgent=True)
            self.hold_time = sim.Exponential(delivery_iat).sample()
            yield self.hold(self.hold_time)
            N = int(env.now()/1440)


"""Antenatal care generator. Pregnant women make four visits to the facility for routine checkup and counselling.
Staff nurses do checkup and maintain the log. ANC visits only happen during the OPD hours. It is assumed (WHO Guidelines)
that a pregnant woman will make 4 visits to the PHC -after the first visit- second, third and fourth visits are scheduled
in next 14, 6 and 6 weeks respectively"""


class ANC(sim.Component):

    global ANC_iat

    env = sim.Environment()
    No_of_shifts = 0                            # tracks number of shifts completed during the simulation time
    No_of_days = 0
    ANC_List = {}
    anc_count = 0
    ANC_p_count = 0

    def process(self):

        sim_time = 0                            # local variable defined for dividing each day into shits
        z = 0
        k = 0
        while z % 3 == 0:                       # condition to run simulation for 8 hour shifts
            ANC.No_of_days += 1                 # class variable to track number of days passed
            while sim_time < 480:               # from morning 8 am to afternoon 4 pm (in minutes)
                ANC.anc_count += 1              # counts overall patients throghout simulation
                ANC.ANC_p_count += 1                 # counts patients in each replication
                id = ANC.anc_count
                age = 223
                day_of_registration = ANC.No_of_days
                visit = 1
                x0 = round(env.now())
                x1 = x0 + 14*7*24*60
                x2 = x0 + 20*7*24*60
                x3 = x0 + 24*7*24*60
                scheduled_visits = [[0], [0], [0], [0]]
                scheduled_visits[0] = x0
                scheduled_visits[1] = x1
                scheduled_visits[2] = x2
                scheduled_visits[3] = x3
                dic = {"ID" : id, "Age": age, "Visit Number" : visit, "Registration day": day_of_registration,
                       "Scheduled Visit": scheduled_visits}
                ANC.ANC_List[id] = dic
                ANC_Checkup()
                ANC_followup(at = ANC.ANC_List[id]["Scheduled Visit"][1])
                ANC_followup(at = ANC.ANC_List[id]["Scheduled Visit"][2])
                ANC_followup(at = ANC.ANC_List[id]["Scheduled Visit"][3])
                hold_time = sim.Exponential(ANC_iat).sample()
                yield self.hold(hold_time)
                sim_time += hold_time
            z += 3
            sim_time = 0
            ANC.No_of_shifts += 3
            yield self.hold(960)                                 # holds simulation for 2 shifts


class ANC_Checkup(sim.Component):

    anc_checkup_count = 0

    def process(self):
        if env.now() <= Main.warm_up:
            self.enter(waitingline_staff_nurse)
            yield self.request(staff_nurse)
            self.leave(waitingline_staff_nurse)
            temp = sim.Triangular(8, 20.6, 12.3).sample()
            yield self.hold(temp)                                   # time taken from a study on ANC visits in Tanzania
            self.enter(waitingline_lab)
            yield self.request(lab)
            self.leave(waitingline_lab)
            yield self.hold(sim.Normal(5,1).bounded_sample())

        else:
            ANC_Checkup.anc_checkup_count += 1
            self.enter(waitingline_staff_nurse)
            yield self.request(staff_nurse)
            self.leave(waitingline_staff_nurse)
            z0 = env.now()
            temp = sim.Triangular(8, 20.6, 12.3).sample()
            yield self.hold(temp)                                   # time taken from a study on ANC visits in Tanzania
            z1 = env.now()
            z = z1 - z0
            Main.SN_time += z
            Main.staff_nurse_ANC += temp
            Main.NT_list.append(temp)
            self.enter(waitingline_lab)
            yield self.request(lab)
            self.leave(waitingline_lab)
            y0 = env.now()
            yield self.hold(sim.Normal(5,1).bounded_sample())
            y1 = env.now()
            y = y1 - y0
            Patient.lab_time += y
            Patient.Lab_time.append(y)


class ANC_followup(sim.Component):

    followup_count = 0

    def process(self):

        if env.now() <= Main.warm_up:
            for key in ANC.ANC_List:                                # for identifying and updating ANC visit number
                x0 = env.now()
                x1 = ANC.ANC_List[key]["Scheduled Visit"][1]
                x2 = ANC.ANC_List[key]["Scheduled Visit"][2]
                x3 = ANC.ANC_List[key]["Scheduled Visit"][3]
                if 0 <= (x1-x0) < 481:
                    ANC.ANC_List[key]["Scheduled Visit"][1] = float("inf")
                    ANC.ANC_List[key]["Visit Number"] = 2
                elif 0 <= (x2-x0) < 481:
                    ANC.ANC_List[key]["Scheduled Visit"][2] = float("inf")
                    ANC.ANC_List[key]["Visit Number"] = 3
                elif 0 <= (x3-x0) < 481:
                    ANC.ANC_List[key]["Scheduled Visit"][3] = float("inf")
                    ANC.ANC_List[key]["Visit Number"] = 4

            self.enter(waitingline_staff_nurse)
            yield self.request(staff_nurse)
            self.leave(waitingline_staff_nurse)                     # calculating staff nurse occupancy
            temp = sim.Triangular(3.33, 13.16, 6.50).sample()
            yield self.hold(temp)
            self.release(staff_nurse)
            self.enter(waitingline_lab)
            yield self.request(lab)
            self.leave(waitingline_lab)
            yield self.hold(sim.Normal(3.456, .823).bounded_sample(2))

        else:
            for key in ANC.ANC_List:                                # for identifying and updating ANC visit number
                x0 = env.now()
                x1 = ANC.ANC_List[key]["Scheduled Visit"][1]
                x2 = ANC.ANC_List[key]["Scheduled Visit"][2]
                x3 = ANC.ANC_List[key]["Scheduled Visit"][3]
                if 0 <= (x1-x0) < 481:
                    ANC.ANC_List[key]["Scheduled Visit"][1] = float("inf")
                    ANC.ANC_List[key]["Visit Number"] = 2
                elif 0 <= (x2-x0) < 481:
                    ANC.ANC_List[key]["Scheduled Visit"][2] = float("inf")
                    ANC.ANC_List[key]["Visit Number"] = 3
                elif 0 <= (x3-x0) < 481:
                    ANC.ANC_List[key]["Scheduled Visit"][3] = float("inf")
                    ANC.ANC_List[key]["Visit Number"] = 4

            ANC_followup.followup_count += 1
            self.enter(waitingline_staff_nurse)
            yield self.request(staff_nurse)
            z0 = env.now()
            self.leave(waitingline_staff_nurse)                     # calculating staff nurse occupancy
            temp = sim.Triangular(3.33, 13.16, 6.50).sample()
            yield self.hold(temp)
            Main.NT_list.append(temp)
            self.release(staff_nurse)
            z1 = env.now()
            z = z1 - z0
            Main.SN_time += z
            Main.staff_nurse_ANC += temp
            self.enter(waitingline_lab)
            yield self.request(lab)
            self.leave(waitingline_lab)
            y0 = env.now()
            yield self.hold(sim.Normal(3.456, .823).bounded_sample(2))
            y1 = env.now()
            y = y1 - y0
            Patient.lab_time += y                                   # calculating lab occupancy
            Patient.Lab_time.append(y)


class Patient(sim.Component):

    global an_list
    global pharm_mean
    global pharm_sd
    global lab_patients
    global ncd_time

    thirty_plus_patients = 0
    OPD_visits = 0
    doc_service_time = []
    doctor_OPD_time = 0
    NCD_Nurse_time_list = []
    Lab_time = []
    lab_time = 0
    pharmacist_time = []
    NCD_Nusre_1_time = 0
    lab_patients = 0

    def process(self):
        """ Includes number of times patients go to OPD, it includes - single, double and triple visits patients.
        Can be less than the actual visits because some visits are scheduled in future after the simulation time. """

        global lab_patients
        global ncd_time

        if env.now() <= Main.warm_up:
            #self.priority = 2       # assigns priority, less than IPD patients
            if OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Age"] >= 30:  #patients's age>30, check BP
                yield self.request((NCD_Nurse, 1))          # requests one staff nurse
                yield self.hold(sim.Uniform(2, 5, 'minutes').sample())             #bounded variable-cannot take negative values
                self.release()
            self.enter(waitingline_OPD)
            yield self.request((doctor,1))
            self.leave(waitingline_OPD)
            consultation_time = sim.Normal(mean, sd, 'minutes').bounded_sample(0.5)
            yield self.hold(consultation_time)
            self.release()
            # for lab visits
            if OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Lab"] == "True":  # checks if lab test is required
                lab_patients += 1
                self.enter(waitingline_lab)
                yield self.request(lab)
                self.leave(waitingline_lab)
                f = round(sim.Normal(3.456, .823).bounded_sample(2), 2)
                yield self.hold(f)
                self.release()
            self.enter(waitingline_pharmacy)
            yield self.request(pharmacist)
            self.leave(waitingline_pharmacy)
            yield self.hold(sim.Normal(pharm_mean, pharm_sd).bounded_sample(.67))
            self.release()

        else:

            Patient.OPD_visits += 1
            """ Includes number of times patients go to OPD, it includes - single, double and triple visits patients.
            Can be less than the actual visits because some visits are scheduled in future after the simulation time. """

            #self.priority = 2       # assigns priority, less than IPD patients
            if OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Age"] >= 30:  #patients's age>30, check BP
                Patient.thirty_plus_patients += 1
                x0 = env.now()
                yield self.request((NCD_Nurse, 1))          # requests one staff nurse
                yield self.hold(sim.Uniform(2, 5, 'minutes').sample())             #bounded variable-cannot take negative values
                self.release()
                x1 = env.now()
                x = (x1-x0)
                Patient.NCD_Nurse_time_list.append(x)
                Patient.NCD_Nusre_1_time += x
                ncd_time += x
            self.enter(waitingline_OPD)
            temp1 = Patient.enter_time(self, waitingline_OPD)
            yield self.request((doctor,1))
            self.leave(waitingline_OPD)
            an_list.append((env.now()-temp1))
            consultation_time = sim.Normal(mean, sd, 'minutes').bounded_sample(0.3)
            yield self.hold(consultation_time)
            time_out = env.now()
            total = round(consultation_time,2)
            Patient.doctor_OPD_time += total
            Patient.doc_service_time.append(total)      # contains all doctor time
            self.release()
            # for lab visits
            if OPD_PatientGenerator.OPD_List[OPD_PatientGenerator.patient_count]["Lab"] == "True":  # checks if lab test is required
                lab_patients += 1
                self.enter(waitingline_lab)
                yield self.request(lab)
                self.leave(waitingline_lab)
                y0 = env.now()
                f = round(sim.Normal(3.456, .823).bounded_sample(2), 2)
                yield self.hold(f)
                self.release()
                y1 = env.now()
                y = y1 - y0
                Patient.lab_time += y
                Patient.Lab_time.append(y)
            self.enter(waitingline_pharmacy)
            yield self.request(pharmacist)
            self.leave(waitingline_pharmacy)
            z0 = env.now()
            yield self.hold(sim.Normal(pharm_mean, pharm_sd).bounded_sample(.67))       #
            self.release()
            z1 = env.now()
            z = z1-z0
            Patient.pharmacist_time.append(z)


class IPD_with_doc(sim.Component):

    global an_list
    global fail_count
    global bed_time

    fail_count =0
    doc_IPD_time = 0                                                # service time of doctor per patient

    def process(self):
        global fail_count
        global bed_time

        if env.now() <= Main.warm_up:
            self.enter_at_head(waitingline_OPD)
            yield self.request(doctor, fail_delay = 20)
            doc_time = round(sim.Uniform(10, 30, 'minutes').sample())
            if self.failed():
                self.leave(waitingline_OPD)
                yield self.request(staff_nurse)
                temp = sim.Uniform(30, 60, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                yield self.request(doctor)
                yield self.hold(doc_time)
                self.release(doctor)
                yield self.request(bed)
                yield self.hold(sim.Triangular(60, 360, 180, 'minutes').bounded_sample(0))
                self.release(bed)
            else:
                self.leave(waitingline_OPD)
                yield self.hold(doc_time)
                self.release(doctor)
                yield self.request(bed, staff_nurse)
                temp = sim.Uniform(30, 60, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                yield self.hold(sim.Triangular(60, 360, 180, 'minutes').bounded_sample(0))
                self.release(bed)

        else:
            self.enter_at_head(waitingline_OPD)
            temp1 = self.enter_time(waitingline_OPD)  # manually calculate queue enter time
            yield self.request(doctor, fail_delay = 20)
            doc_time = round(sim.Uniform(10, 30, 'minutes').sample())
            Patient.doc_service_time.append(doc_time)   # contains all doctor time
            IPD_with_doc.doc_IPD_time += doc_time

            if self.failed():
                self.leave(waitingline_OPD)
                an_list.append((env.now()-temp1))           # list that contains queue waiting time for del patients
                fail_count +=1
                yield self.request(staff_nurse)
                z0 = env.now()
                temp = sim.Uniform(30, 60, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                yield self.request(doctor)
                yield self.hold(doc_time)
                self.release(doctor)
                Main.NT_list.append(temp)                   # stores nurse time in the list
                z1 = env.now()
                z = z1 - z0
                Main.SN_time += z
                Main.staff_nurse_IPD += temp
                yield self.request(bed)
                y = sim.Triangular(60, 360, 180, 'minutes').bounded_sample(0)
                bed_time += y
                yield self.hold(y)
                self.release(bed)

            else:
                self.leave(waitingline_OPD)
                an_list.append((env.now()-temp1))
                yield self.hold(doc_time)

                self.release(doctor)
                yield self.request(bed, staff_nurse)
                z0 = env.now()
                temp = sim.Uniform(30, 60, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                Main.NT_list.append(temp)                   # stores nurse time in the list
                z1 = env.now()
                z = z1 - z0
                Main.SN_time += z
                Main.staff_nurse_IPD += temp
                y= sim.Triangular(60, 360, 180, 'minutes').bounded_sample(0)
                yield self.hold(y)
                bed_time += y
                self.release(bed)


class IPD_no_doc(sim.Component):                                # during night when no doctor is available

    global bed_time

    def process(self):
        global bed_time

        if env.now() <= Main.warm_up:
            yield self.request(bed, staff_nurse)
            temp = sim.Uniform(30, 60, 'minutes').sample()
            yield self.hold(temp)
            self.release(staff_nurse)
            yield self.hold(sim.Triangular(60, 360, 180, 'minutes').sample())
            self.release(bed)
        else:
            yield self.request(bed, staff_nurse)
            z0 = env.now()
            temp = sim.Uniform(30, 60, 'minutes').sample()
            yield self.hold(temp)
            self.release(staff_nurse)
            z1 = env.now()
            z = z1 - z0
            Main.SN_time += z
            Main.staff_nurse_IPD += temp
            Main.NT_list.append(temp)
            y = (sim.Triangular(60, 360, 180, 'minutes').sample())
            bed_time += y
            yield self.hold(y)
            self.release(bed)


class Delivery_with_doctor(sim.Component):

    global an_list
    global fail1

    doc_delivery_time = 0
    del_OPD = 0
    an_list = []
    fail1 = 0

    def process(self):
        global fail1
        global fail_count
        global bed_time

        if env.now() <= Main.warm_up:
            self.enter_at_head(waitingline_OPD)
            yield self.request((doctor, 1), fail_delay=20)
            doc_time = round(sim.Uniform(30, 60, 'minutes').sample(),2)
            if self.failed():
                self.leave(waitingline_OPD)
                yield self.request(staff_nurse)
                temp = sim.Uniform(120, 240, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                yield self.request(doctor)
                yield self.hold(doc_time)
                self.release(doctor)
                yield self.request(delivery_bed, fail_delay=120)
                if self.failed():
                    pass
                else:
                    yield self.hold(sim.Uniform(6*60, 10*60, 'minutes').sample())
                    self.release(delivery_bed)
                    yield self.request(bed)
                    yield self.hold(sim.Uniform(240, 1440, 'minutes').sample())
            else:
                self.leave(waitingline_OPD)
                yield self.hold(doc_time)
                self.release(doctor)
                yield self.request(staff_nurse)
                temp = sim.Uniform(120,240, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                yield self.request(delivery_bed)
                yield self.hold(sim.Uniform(6*60, 10*60, 'minutes').sample())
                self.release(delivery_bed)
                yield self.request(bed)
                yield self.hold(sim.Uniform(240, 1440, 'minutes').bounded_sample(0))     # holding patient for min 4 hours
                # to 48 hours
                self.release(bed)

        else:
            Delivery_with_doctor.del_OPD += 1
            self.enter_at_head(waitingline_OPD)
            temp1 = self.enter_time(waitingline_OPD)  # manually calculate queue enter time
            yield self.request((doctor, 1), fail_delay=20)
            doc_time = round(sim.Uniform(30, 60, 'minutes').sample(),2)
            Patient.doc_service_time.append(doc_time)
            Delivery_with_doctor.doc_delivery_time += doc_time

            if self.failed():
                fail1 +=1
                self.leave(waitingline_OPD)
                an_list.append((env.now()-temp1))          # list that contains queue waiting time for del patients
                yield self.request(staff_nurse)
                z0 = env.now()
                temp = sim.Uniform(120, 240, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                Main.NT_list.append(temp)
                z1 = env.now()
                z = z1 - z0
                Main.SN_time += z
                Main.staff_nurse_del += temp
                yield self.request(doctor)
                yield self.hold(doc_time)

                self.release(doctor)
                yield self.request(delivery_bed, fail_delay=120)
                if self.failed():
                    fail_count += 1
                    Delivery_with_doctor.del_OPD -= 1
                else:
                    yield self.hold(sim.Uniform(6*60, 10*60, 'minutes').sample())
                    self.release(delivery_bed)
                    yield self.request(bed)
                    y = sim.Uniform(240, 1440, 'minutes').sample()
                    bed_time += y
                    yield self.hold(y)
            else:
                self.leave(waitingline_OPD)
                an_list.append((env.now()-temp1))          # list that contains queue waiting time for del patients
                yield self.hold(doc_time)
                self.release(doctor)
                yield self.request(staff_nurse)
                z0 = env.now()
                temp = sim.Uniform(120,240, 'minutes').sample()
                yield self.hold(temp)
                self.release(staff_nurse)
                Main.NT_list.append(temp)
                z1 = env.now()
                z = z1 - z0
                Main.SN_time += z
                Main.staff_nurse_del += temp
                yield self.request(delivery_bed)
                yield self.hold(sim.Uniform(6*60, 10*60, 'minutes').sample())
                self.release(delivery_bed)
                yield self.request(bed)
                y = sim.Uniform(240, 1440, 'minutes').sample()
                bed_time += y
                yield self.hold(y)     # holding patient for min 4 hours
                # to 48 hours
                self.release(bed)


class Delivery_no_doc(sim.Component):                                       # during night when no doctor is available

    del_after_OPD = 0

    def process(self):

        global fail_count
        global bed_time

        if env.now() <= Main.warm_up:
            yield self.request(staff_nurse)
            temp = sim.Uniform(120, 240, 'minutes').sample()
            yield self.hold(temp)
            self.release(staff_nurse)
            yield self.request(delivery_bed, fail_delay=120)
            if self.failed():
                pass
            else:
                yield self.hold(sim.Uniform(6*60, 10*60, 'minutes').sample())
                self.release(delivery_bed)
                yield self.request(bed)
                yield self.hold(sim.Uniform(240, 1440, 'minutes').bounded_sample(0))
                self.release(bed)

        else:
            Delivery_no_doc.del_after_OPD += 1
            yield self.request(staff_nurse)
            z0 = env.now()
            temp = sim.Uniform(120, 240, 'minutes').sample()
            yield self.hold(temp)
            self.release(staff_nurse)
            Main.NT_list.append(temp)
            z1 = env.now()
            z = z1 - z0
            Main.SN_time += z
            Main.staff_nurse_del += temp
            yield self.request(delivery_bed, fail_delay=120)
            if self.failed():
                fail_count += 1
                Delivery_no_doc.del_after_OPD -= 1
            else:
                yield self.hold(sim.Uniform(6*60, 10*60, 'minutes').sample())
                self.release(delivery_bed)
                yield self.request(bed)
                y = sim.Uniform(240, 1440, 'minutes').sample()
                bed_time += y
                yield self.hold(y)
                self.release(bed)


def main():

    # defining simulation input parameters
    global OPD_iat
    global delivery_iat
    global IPD_iat
    global ANC_iat
    global mean
    global sd
    global pharm_mean
    global pharm_sd
    global j
    global sumi
    global days
    global shifts
    global hours
    global doc_cap
    global staff_nurse_cap
    global NCD_nurse_cap
    global pharmacist_cap
    global lab_cap

    # defining salabim resources
    global env
    global doctor
    global staff_nurse
    global NCD_Nurse
    global pharmacist
    global lab
    global bed
    global delivery_bed

    # Defining salabim queues
    global waitingline_OPD
    global waitingline_staff_nurse
    global waitingline_pharmacy
    global waitingline_lab

    # defining lists and variables for collecting data
    global lisst
    global OPD_q_waiting_time_list
    global OPD_q_length_list
    global pharmacy_q_waiting_time_list
    global lab_q_waiting_time_list
    global pharmacy_q_length_list
    global lab_q_length_list
    global lab_patient_list
    global OPD_q__list

    global OPD_patients_list         # stores no of out-patients generated per replication
    global IPD_patients_list         # stores no of in-patients generated per replication
    global ANC_patients_list         # stores no of ANC-patients generated per replication
    global Delivery_patients_list    # stores no of delivery patients generated per replication
    global bed_occupancy_list
    global doc_occupancy
    global doc_tot_time
    global lab_patients
    global an_list
    global NCD_occ_list
    global lab_occ_list
    global pharm_occ_list
    global nurse_occ_list
    global total_admin_work          # admin work per replication
    global delivery_bed_occ_list
    global fail_count_list
    global fail_count

    global f
    global f1
    global f2
    global f3
    global f4
    global f5

    global ncd_time
    global ncd_util

    global bed_time


    ncd_util =[]
    bed_util = []

    OPD_q_waiting_time_list = []            # stores waiting time in OPD queue for each replication
    OPD_q_length_list = []                  # stores avg length of OPD queue for each replication
    pharmacy_q_waiting_time_list = []
    pharmacy_q_length_list = []
    lab_q_waiting_time_list = []
    lab_q_length_list = []
    OPD_patients_list = []            # to store no of patients generated in each replication
    IPD_patients_list = []
    ANC_patients_list = []
    Delivery_patients_list = []
    wait_time_OPD = []
    doc_occupancy = []
    lab_patient_list = []
    OPD_q__list = []
    bed_occupancy_list = []
    NCD_occ_list = []
    total_admin_work = []
    lab_occ_list = []
    pharm_occ_list = []
    nurse_occ_list = []
    delivery_bed_occ_list =[]
    fail_count_list = []

    OPD_iat = 4
    delivery_iat = 1440              # inter-arrival delivery patient time
    IPD_iat = 2880                   # inter-arrival IPD patient time
    ANC_iat = 1440                   # inter-arrival ANC patient time
    mean = 0.87                    # consultation time mean
    sd = .21                        # consultation time sd
    pharm_mean = 2.083
    pharm_sd = 0.72
    j = 0
    f = 0                            # for calculating sum of OPD q waiting time
    f1 = 0                           # for calculating sum of OPD q length
    f2 = 0                           # for calculating sum of pharmacy q waiting time
    f3 = 0                           # for calculating sum of pharmacy q length
    f4 = 0                           # for calculating sum of lab q waiting time
    f5 = 0                           # for calculating sum of lab q length

    doc_tot_time = 0
    lab_patients = 0

    days = 365
    shifts = 3
    hours = 8
    doc_cap = 2                      # number of doctors
    staff_nurse_cap = 3              # number of nurses
    NCD_nurse_cap = 1                # number of NCD nurses
    pharmacist_cap = 1               # number of pharmacists
    lab_cap = 1                      # number of lab technicians
    replication = 10

    for x in range(0, replication):
        n = np.random.randint(0, 101)
        env = sim.Environment(trace=False, random_seed='', time_unit='minutes')
        simulation_time = days*shifts*hours*60
        NCD_Nurse = sim.Resource("Staff nurse 1", capacity=NCD_nurse_cap )
        staff_nurse = sim.Resource("Staff nurse", capacity=3)
        doctor = sim.Resource('doctor', capacity = doc_cap)
        lab = sim.Resource('Lab', capacity=lab_cap)
        pharmacist = sim.Resource("Pharmacy", capacity=pharmacist_cap)
        bed = sim.Resource("Bed", capacity=6)
        delivery_bed = sim.Resource("Del bed", capacity = 1)
        Main(name='')
        IPD_PatientGenerator(name="IPD_Patient")
        Delivery(name="Delivery Patient")
        ANC(name="ANC Patients")
        waitingline_staff_nurse = sim.Queue("waitingline_staff_nurse 1")
        waitingline_OPD = sim.Queue('waitingline_OPD')
        waitingline_lab = sim.Queue("waitingline_lab")
        waitingline_pharmacy = sim.Queue("waitingline_pharmacy")
        doc_tot_time = 0                # for each replication makes this zero
        Patient.doc_service_time = []   # at the start of replication makes this array empty
        bed_time = 0

        """Actual simulation"""
        env.run(till=(simulation_time+ Main.warm_up))
        print(" No of replications done", x)

        sumi = 0
        for i in Main.NCD_admin_work:
            sumi += i
        Main.NCD_admin_work = []        # empties admin work
        total_admin_work.append(sumi)
        j = 0
        for i in Patient.NCD_Nurse_time_list:
            j += i
        Total_NCD_time = 0
        Total_NCD_time = j + total_admin_work[x]
        k = 0
        for i in Patient.Lab_time:
            k += i
        ncd_util.append(ncd_time/(420*365))
        ncd_time=0

        bed_util.append(bed_time/(days*1440*6))

        """ To update lists containing waiting time, number of patients and average length per replication"""
        OPD_q_waiting_time_list.append(waitingline_OPD.length_of_stay.mean())
        OPD_q_length_list.append(waitingline_OPD.length.mean())
        pharmacy_q_waiting_time_list.append(waitingline_pharmacy.length_of_stay.mean())
        pharmacy_q_length_list.append(waitingline_pharmacy.length.mean())
        lab_q_waiting_time_list.append(waitingline_lab.length_of_stay.mean())
        lab_q_length_list.append(waitingline_lab.length.mean())
        wait_time_OPD.append(np.mean(waitingline_OPD.length_of_stay.x()))
        fail_count_list.append(fail_count)
        fail_count = 0

        """Monitoring list of patients for every replication """
        OPD_patients_list.append(Main.p_count)  # gives no of patients for each replication
        ANC_patients_list.append(ANC.ANC_p_count)
        IPD_patients_list.append(IPD_PatientGenerator.p_count)
        Delivery_patients_list.append(Delivery.p_count)
        bed_occupancy_list.append(round(bed.occupancy.mean(), 2))
        delivery_bed_occ_list.append(round(delivery_bed.occupancy.mean(), 2))
        Main.p_count = 0
        ANC.ANC_p_count = 0
        Delivery.p_count = 0
        IPD_PatientGenerator.p_count = 0
        "  occupancy for each replication"
        doc_tot_time = sum(Patient.doc_service_time)        # cal total doc time during 365 days
        z = doc_tot_time/(420*days*doc_cap)                  # cal the occupancy for eac replication
        doc_occupancy.append(z)                             # stores occupancy in an array
        lab_patient_list.append(lab_patients)
        OPD_q__list.append(np.mean(an_list))
        an_list = []
        NCD_occ_list.append((Total_NCD_time) / (420 * days))
        lab_occ_list.append(k/(420*days))
        pharm_occ_list.append(sum(Patient.pharmacist_time) / (420 * days))
        nurse_occ_list.append((sum(Main.NT_list)+180*days)/(21*60*days))
        Main.NT_list = []
        Patient.pharmacist_time = []
        Patient.Lab_time = []
        Patient.NCD_Nurse_time_list = []
    j = 0
    row = 1
    col = 1
    ad = 0
    temp = 0

    REPLICATION = xlsxwriter.Workbook("Config_3(2).xlsx")
    worksheet = REPLICATION.add_worksheet("Sheet 3")

    # input parameters
    worksheet.write(0, 0, "OPD patient inter-arrival time")
    worksheet.write(row, 0, "IPD patient inter-arrival time")
    worksheet.write(row+1, 0, "Delivery patient inter-arrival time")
    worksheet.write(row+2, 0, "ANC patient inter-arrival time")
    worksheet.write(row+3, 0, "Doctor consultation time mean")
    worksheet.write(row+4, 0, "Doctor consultation time SD")

    # Resources
    worksheet.write(row+5, 0, "Number of doctor")
    worksheet.write(row+6, 0, "Number of staff nurses")
    worksheet.write(row+7, 0, "Number of lab technician")
    worksheet.write(row+8, 0, "Number of pharmacist")

    # Time outputs
    worksheet.write(row+9, 0, "Total OPD minutes")
    worksheet.write(row+10, 0, "OPD patients per day")
    worksheet.write(row+11, 0, "OPD patients per month")
    worksheet.write(row+12, 0, "Doctor OPD time")
    worksheet.write(row+13, 0, "Doctor IPD time")
    worksheet.write(row+14, 0, "Doctor Delivery time")
    worksheet.write(row+15, 0, "Doctor admin time")
    worksheet.write(row+16, 0, "Doctor total time")
    worksheet.write(row+17, 0, "NCD Nurse OPD time")
    worksheet.write(row+18, 0, "NCD Nurse Admin time")
    worksheet.write(row+19, 0, "Staff Nurse IPD time")
    worksheet.write(row+20, 0, "Staff Nurse Delivery time")
    worksheet.write(row+21, 0, "Staff Nurse ANC time")
    worksheet.write(row+22, 0, "Staff Nurse Admin time")
    worksheet.write(row+23, 0, "Staff Nurse total ime")
    worksheet.write(row+24, 0, "Pharmacist Time")
    worksheet.write(row+25, 0, "Lab Time")
    worksheet.write(row+26, 0, "Average time spent with doctors")
    worksheet.write(row+27, 0, "Average time of bed occupancy")

    # Occupancy
    worksheet.write(row+28, 0, "Doctor Occupancy")
    worksheet.write(row+29, 0, "NCD Nurse Occupancy")
    worksheet.write(row+30, 0, "Staff nurse Occupancy")
    worksheet.write(row+31, 0, "Pharmacist Occupancy")
    worksheet.write(row+32, 0, "Lab Occupancy")
    worksheet.write(row+33, 0, "Bed occupancy")

    # Queue
    worksheet.write(row+34, 0, "Mean length of OPD queue")
    worksheet.write(row+35, 0, "OPD queue waiting time")
    worksheet.write(row+36, 0, "Mean length of pharmacy queue")
    worksheet.write(row+37, 0, "Pharmacy queue waiting time")
    worksheet.write(row+38, 0, "Mean length of Lab queue")
    worksheet.write(row+39, 0, "Lab queue waiting time")

    # Numbers
    worksheet.write(row+40, 0, "OPDs")
    worksheet.write(row+41, 0, "IPDs")
    worksheet.write(row+42, 0, "Deliveries")
    worksheet.write(row+43, 0, "ANCs")
    worksheet.write(row+44, 0, "Total ANC visits")
    worksheet.write(row+45, 0, "Delivery in OPD time")
    worksheet.write(row+46, 0, "Delivery after OPD time")
    worksheet.write(row+47, 0, "Replications")

    # Stats
    # standard deviations and max
    worksheet.write(row+48, 0, "OPD queue wait time SD")
    worksheet.write(row+49, 0, "OPD queue max wait time")
    worksheet.write(row+50, 0, "Pharmacy queue wait time SD")
    worksheet.write(row+51, 0, "Pharmacy queue max wait time")
    worksheet.write(row+52, 0, "lab queue wait time SD")
    worksheet.write(row+53, 0, "Lab queue max wait time")
    worksheet.write(row+54, 0, "Lab Patients")
    worksheet.write(row+55, 0, "Lab patients std")

    # Outputs
    """Input parameters"""
    worksheet.write(row - 1, col, OPD_iat)
    worksheet.write(row, col, IPD_iat)
    worksheet.write(row+1, col, delivery_iat)
    worksheet.write(row+2, col, ANC_iat)
    worksheet.write(row+3, col, mean)
    worksheet.write(row+4, col, sd)

    """Resources"""
    worksheet.write(row+5, col, doc_cap)
    worksheet.write(row+6, col, staff_nurse_cap)
    worksheet.write(row+7, col, lab_cap)
    worksheet.write(row+8, col, pharmacist_cap)

    # output time
    """Doctor time calculation"""

    # OPD minutes
    worksheet.write(row+9, col, (Main.No_of_days*480)/replication)

    # OPD patients per day
    worksheet.write(row+10, col, round(Patient.OPD_visits/days)/replication)

    # OPD patients per month
    worksheet.write(row+11, col, round((Patient.OPD_visits)/Main.No_of_days)*26)

    # Doctor OPD time
    worksheet.write(row+12, col, round(Patient.doctor_OPD_time/replication, 2))

    # Doctor IPD time
    worksheet.write(row+13, col, round(IPD_with_doc.doc_IPD_time/replication, 2))

    # Doctor Delivery time
    worksheet.write(row+14, col, round(Delivery_with_doctor.doc_delivery_time/replication, 2))

    # Doctor admin time
    worksheet.write(row+15, col, sumi/replication)

    # Total doctor time
    worksheet.write_formula(row+16, col, '=sum(B17+B14+B16+B15)')

    """NCD time calculation"""
    for i in Patient.NCD_Nurse_time_list:
        j += i
    # NCD OPD time
    worksheet.write(row+17, col, round(j, 2)/replication)
    worksheet.write(row+18, col, sumi/replication)

    """Staff Nurse time calculation"""
    worksheet.write(row+19, col,  round(Main.staff_nurse_IPD, 2)/replication)
    worksheet.write(row+20, col,  round(Main.staff_nurse_del, 2)/replication)
    worksheet.write(row+21, col,  round(Main.staff_nurse_ANC,2)/replication)

    ad = round(random.normalvariate(60, 10), 2)                    # admin time
    worksheet.write(row+22, col, ad*3)
    worksheet.write(row+23, col,  '=sum(B20+B21+B22+B23)')

    """Pharmacy time calculation"""
    tim = 0
    for time in Patient.pharmacist_time:
        tim += time
    worksheet.write(row+24, col, round(tim,2)/replication)

    """Lab time calculation"""
    worksheet.write(row+25, col, round(Patient.lab_time,)/replication)

    # average consultation time
    worksheet.write(row+26, col, round(doctor.claimers().length_of_stay.mean()))
    worksheet.write(row+27, col, round(bed.claimers().length_of_stay.mean()))

    """Occupancy"""
    # Doctor
    worksheet.write(row+28, col, (sum(doc_occupancy))/replication)
    # NCD
    mean1 = sum(NCD_occ_list)
    worksheet.write(row+29, col, round(mean1/replication, 2))

    # Staff Nurse
    mean = sum(nurse_occ_list)
    worksheet.write(row+30, col, round(mean/replication, 2))

    # Pharmacy
    worksheet.write(row+31, col, sum(pharm_occ_list)/replication)

    # Lab
    worksheet.write(row+32, col, round(Patient.lab_time/(420*days*replication), 2))

    # Bed
    worksheet.write(row+33, col, round(bed.occupancy.mean(), 2))

    """Queue outputs"""
    for _ in OPD_q_waiting_time_list:
        f = f+_
    for _ in OPD_q_length_list:
        f1 += _
    for _ in pharmacy_q_waiting_time_list:
        f2 += _
    for _ in pharmacy_q_length_list:
        f3 += _
    for _ in lab_q_waiting_time_list:
        f4 += _
    for _ in lab_q_length_list:
        f5 += _

    # OPD queue average length
    worksheet.write(row+34, col, round(f1/replication, 4))
    # OPD queue waiting time
    worksheet.write(row+35, col, round(f/replication, 4))

    # Pharmacy
    # Pharmacy queue length
    worksheet.write(row+36, col, round(f3/replication, 2))
    # Pharmacy queue waiting time
    worksheet.write(row+37, col, round(f2/replication, 2))

    # Lab
    # lab queue length
    worksheet.write(row+38, col, round(f5/replication, 2))
    # lab queue waiting time
    worksheet.write(row+39, col, round(f4/replication, 2))

    """Output Numbers"""
    worksheet.write(row+40, col, Patient.OPD_visits/replication)
    worksheet.write(row+41, col, (len(IPD_PatientGenerator.IPD_List))/replication)
    worksheet.write(row+42, col, Delivery.Delivery_count/replication)
    worksheet.write(row+43, col, (len(ANC.ANC_List))/replication)
    worksheet.write(row+44, col, (ANC_Checkup.anc_checkup_count + ANC_followup.followup_count)/replication)
    worksheet.write(row+45, col, Delivery_with_doctor.del_OPD/replication)
    worksheet.write(row+46, col, Delivery_no_doc.del_after_OPD/replication)
    worksheet.write(row+47, col, replication)
    worksheet.write(row+48, col, statistics.stdev(OPD_q_waiting_time_list))
    worksheet.write(row+49, col, max(OPD_q_waiting_time_list))
    worksheet.write(row+50, col, statistics.stdev(pharmacy_q_waiting_time_list))
    worksheet.write(row+51, col, max(pharmacy_q_waiting_time_list))
    worksheet.write(row+52, col, statistics.stdev(lab_q_waiting_time_list))
    worksheet.write(row+53, col, max(lab_q_waiting_time_list))
    worksheet.write(row+54, col, sum(lab_patient_list)/replication)
    worksheet.write(row+55, col, np.std(lab_patient_list))
    #REPLICATION.close()

    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet 1")
    ws.write(0, 0, "OPD patients")
    ws.write(1, 0, "IPD patients")
    ws.write(2, 0, "ANC patients")
    ws.write(3, 0, "Del patients")
    ws.write(4, 0, "OPD Q wt")
    ws.write(5, 0, "Pharmacy Q wt")
    ws.write(6, 0, "Lab Q wt")
    ws.write(7, 0, "doc occ")
    ws.write(8, 0, "Lab patient list")
    ws.write(9, 0, "OPD q len")
    ws.write(10, 0, "ipd occ")
    ws.write(11, 0, "opd q len")
    ws.write(12, 0, "pharmacy q len")
    ws.write(13, 0, "lab q len")
    ws.write(14, 0, "NCD occ")
    ws.write(15, 0, "lab occ")
    ws.write(16, 0, "pharm occ")
    ws.write(17, 0, "staff nurse occ")
    ws.write(18, 0, "del occ")
    ws.write(19, 0, "del referred")
    ws.write(20, 0, "NCD occ")
    ws.write(21, 0, "ipd bed occ")

    for index, item in enumerate(OPD_patients_list):
        ws.write(0, index+1, item)
    for index, item in enumerate(IPD_patients_list):
        ws.write(1, index+1, item)
    for index, item in enumerate(ANC_patients_list):
        ws.write(2, index+1, item)
    for index, item in enumerate(Delivery_patients_list):
        ws.write(3, index+1, item)
    for index, item in enumerate(OPD_q_waiting_time_list):
        ws.write(4, index+1, item)
    for index, item in enumerate(pharmacy_q_waiting_time_list):
        ws.write(5, index+1, item)
    for index, item in enumerate(lab_q_waiting_time_list):
        ws.write(6, index+1, item)
    for index, item in enumerate(doc_occupancy):
        ws.write(7, index+1, item)
    for index, item in enumerate(lab_patient_list):
        ws.write(8, index+1, item)
    for index, item in enumerate(OPD_q__list):
        ws.write(9, index+1, item)
    for index, item in enumerate(bed_occupancy_list):
        ws.write(10, index+1, item)
    for index, item in enumerate(OPD_q_length_list):        # length of OPD queue for each replication
        ws.write(11, index+1, item)
    for index, item in enumerate(pharmacy_q_length_list):   # length of Pharmacy queue for each replication
        ws.write(12, index+1, item)
    for index, item in enumerate(lab_q_length_list):        # length of lab for each replication
        ws.write(13, index+1, item)
    for index, item in enumerate(NCD_occ_list):        # length of lab for each replication
        ws.write(14, index+1, item)
    for index, item in enumerate(lab_occ_list):        # length of lab for each replication
        ws.write(15, index+1, item)
    for index, item in enumerate(pharm_occ_list):        # length of lab for each replication
        ws.write(16, index+1, item)
    for index, item in enumerate(nurse_occ_list):        # length of lab for each replication
        ws.write(17, index+1, item)
    for index, item in enumerate(delivery_bed_occ_list):        # length of lab for each replication
        ws.write(18, index+1, item)
    for index, item in enumerate(fail_count_list):        # length of delivery patients tured away
        ws.write(19, index+1, item)
    for index, item in enumerate(ncd_util):        # length of delivery patients tured away
        ws.write(20, index+1, item)

    for index, item in enumerate(bed_util):        # length of delivery patients tured away
        ws.write(21, index+1, item)

    wb.save('Outputs.xls')


if __name__ == '__main__':
    main()


