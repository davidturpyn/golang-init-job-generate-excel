package models

import "gorm.io/gorm"

type User struct {
	gorm.Model
	Name  string `gorm:"size:100;not null"`
	Email string `gorm:"size:255;unique;not null"`
	Age   int    `gorm:"not null"`
}
