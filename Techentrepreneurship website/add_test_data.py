"""
Script zum Hinzufügen von Testhobbys in die Datenbank
"""
from app import app, db, User, Hobby

def add_test_data():
    with app.app_context():
        # Erstelle einen Test-Benutzer
        test_user = User.query.filter_by(username='testuser').first()
        if not test_user:
            test_user = User(
                username='testuser',
                email='test@rentahobby.ch',
                password='hashed_password_here'
            )
            db.session.add(test_user)
            db.session.commit()
        
        # Teste Hobbys basierend auf HOBBY_TRAITS aus app.py
        hobbies_data = [
            {
                'name': 'Photography',
                'description': 'Lerne die Kunst der Fotografie! Wir stellen dir eine professionelle Kamera zur Verfügung und zeigen dir die besten Techniken zum Fotografieren.',
                'category': 'Kunst & Kreativität',
                'price_per_day': 25.0,
                'image_url': 'https://images.unsplash.com/photo-1606986628025-35d57e735ae0?w=400'
            },
            {
                'name': 'Pottery',
                'description': 'Entdecke deine kreative Seite beim Töpfern! Mit Ton und Drehscheibe entstehen wunderbare Kunstwerke.',
                'category': 'Kunst & Kreativität',
                'price_per_day': 30.0,
                'image_url': 'https://images.unsplash.com/photo-1578926314433-b22ef9cfc999?w=400'
            },
            {
                'name': 'Bowling',
                'description': 'Spaß und Action im Bowling-Center! Perfekt für Freizeit und Wettkampf mit Freunden.',
                'category': 'Sport',
                'price_per_day': 20.0,
                'image_url': 'https://images.unsplash.com/photo-1533900298318-6b8da08a523e?w=400'
            },
            {
                'name': 'Saxophone',
                'description': 'Spiele Jazz, Blues und mehr mit dem Saxophon! Anfänger und Fortgeschrittene willkommen.',
                'category': 'Musik',
                'price_per_day': 35.0,
                'image_url': 'https://images.unsplash.com/photo-1510915361894-db8b60106cb1?w=400'
            },
            {
                'name': 'Painting',
                'description': 'Maler dein Meisterwerk! Mit Acryl, Aquarell oder Öl - wir haben alles was du brauchst.',
                'category': 'Kunst & Kreativität',
                'price_per_day': 28.0,
                'image_url': 'https://images.unsplash.com/photo-1586985289688-face91d3749e?w=400'
            },
            {
                'name': 'Yoga',
                'description': 'Finde innere Ruhe und Ausgeglichenheit durch Yoga. Kurse für alle Fitnesslevel.',
                'category': 'Wellness',
                'price_per_day': 22.0,
                'image_url': 'https://images.unsplash.com/photo-1506126613408-eca07ce68773?w=400'
            },
            {
                'name': 'Guitar',
                'description': 'Erlernen Sie Gitarre spielen! Von klassisch bis Rock - es ist alles möglich.',
                'category': 'Musik',
                'price_per_day': 32.0,
                'image_url': 'https://images.unsplash.com/photo-1423666639041-f56000c27a7a?w=400'
            },
            {
                'name': 'Tennis',
                'description': 'Trainiere deinen Tennisschlag und verbessere dein Spiel! Plätze und Ausrüstung vorhanden.',
                'category': 'Sport',
                'price_per_day': 24.0,
                'image_url': 'https://images.unsplash.com/photo-1554068865-24cecd4e34c8?w=400'
            },
            {
                'name': 'Meditation',
                'description': 'Trainiere dein Gehirn mit regelmäßigen Meditationen. Kurse für Anfänger bis Profis.',
                'category': 'Wellness',
                'price_per_day': 20.0,
                'image_url': 'https://images.unsplash.com/photo-1506126613408-eca07ce68773?w=400'
            },
            {
                'name': 'Dance',
                'description': 'Tanze Jazz, Hip-Hop, Contemporary und mehr! Professionelle Trainer führen dich an.',
                'category': 'Sport & Bewegung',
                'price_per_day': 26.0,
                'image_url': 'https://images.unsplash.com/photo-1509631179647-0177331693ae?w=400'
            }
        ]
        
        # Überprüfe ob Hobbys schon existieren und füge neue hinzu
        for hobby_data in hobbies_data:
            existing = Hobby.query.filter_by(name=hobby_data['name']).first()
            if not existing:
                hobby = Hobby(
                    name=hobby_data['name'],
                    description=hobby_data['description'],
                    category=hobby_data['category'],
                    price_per_day=hobby_data['price_per_day'],
                    image_url=hobby_data['image_url'],
                    user_id=test_user.id
                )
                db.session.add(hobby)
                print(f'✓ Hobby "{hobby_data["name"]}" hinzugefügt')
            else:
                print(f'- Hobby "{hobby_data["name"]}" existiert bereits')
        
        db.session.commit()
        print('\n✓ Testdaten erfolgreich hinzugefügt!')

if __name__ == '__main__':
    add_test_data()
