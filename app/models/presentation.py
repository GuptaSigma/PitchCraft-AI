from app import db
from datetime import datetime

class Presentation(db.Model):
    """Presentation model with theme support"""
    __tablename__ = 'presentations'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, index=True)
    title = db.Column(db.String(500), nullable=False)
    description = db.Column(db.Text, nullable=True)
    content = db.Column(db.JSON, nullable=False)  # Store slides as JSON
    theme = db.Column(db.String(50), default='dialogue')  # CRITICAL: Store selected theme
    style = db.Column(db.String(50), default='business')
    language = db.Column(db.String(10), default='en-uk')
    image_style = db.Column(db.String(50), default='illustration')
    total_slides = db.Column(db.Integer, default=10)
    is_public = db.Column(db.Boolean, default=False)
    views = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        """Convert to dictionary"""
        return {
            'id': self.id,
            'user_id': self.user_id,
            'title': self.title,
            'description': self.description,
            'content': self.content,
            'theme': self.theme,  # Include theme in response
            'style': self.style,
            'language': self.language,
            'image_style': self.image_style,
            'total_slides': self.total_slides,
            'is_public': self.is_public,
            'views': self.views,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }
    
    def __repr__(self):
        return f'<Presentation {self.title} (theme={self.theme})>'