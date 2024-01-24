

def test (self):
    list_start_end_road = self.data.get('Ось дороги', {}).get('Начало трассы', [])
    list_start_end_km_sign = self.data.get('Километровые знаки').get('Значение в прямом направлении')
    l = {}
    # min(enumerate(a), key = lambda x: abs(x[1] - 11.5))
    for num_road in list_start_end_road:
        for num_sign in list_start_end_km_sign:
            if -1000 < num_road[1] - num_sign[1] < 1000:
                print(num_road, 'start',  num_sign)
                #l['start'] += (num_road[1] - num_sign[1])
                if -1000 < num_road[2] - num_sign[2] < 1000:
                    print(num_road, 'end', num_sign)
                    #l['end'] += (num_road[2] - num_sign[2])