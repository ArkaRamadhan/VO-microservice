package models

import (
	"encoding/json"
	"time"
)

// model for meeting
type Meeting struct {
	ID               uint       `gorm:"primaryKey"`
	CreatedAt        *time.Time `gorm:"autoCreateTime"`
	UpdatedAt        *time.Time `gorm:"autoUpdateTime"`
	Task             *string    `json:"task"`
	TindakLanjut     *string    `json:"tindak_lanjut"`
	Status           *string    `json:"status"`
	UpdatePengerjaan *string    `json:"update_pengerjaan"`
	Pic              *string    `json:"pic"`
	TanggalTarget    *time.Time `json:"tanggal_target"`
	TanggalActual    *time.Time `json:"tanggal_actual"`
	CreateBy         string     `json:"create_by"`
}

func (i *Meeting) MarshalJSON() ([]byte, error) {
	type Alias Meeting
	tanggalTargetFormatted := i.TanggalTarget.Format("2006-01-02")
	tanggalActualFormatted := i.TanggalActual.Format("2006-01-02")
	return json.Marshal(&struct {
		TanggalTarget *string `json:"tanggal_target"`
		TanggalActual *string `json:"tanggal_actual"`
		*Alias
	}{
		TanggalTarget: &tanggalTargetFormatted,
		TanggalActual: &tanggalActualFormatted,
		Alias:         (*Alias)(i),
	})
}

type BookingRapat struct {
	ID     uint   `gorm:"primaryKey" json:"id"`
	Title  string `json:"title"`
	Start  string `json:"start"`
	End    string `json:"end"`
	AllDay bool   `json:"allDay"`
	Color  string `json:"color"` // Tambahkan field ini untuk warna
	Status string `json:"status"`
}

func (BookingRapat) TableName() string {
	return "booking_rapats"
}

type JadwalRapat struct {
	ID     uint   `gorm:"primaryKey" json:"id"`
	Title  string `json:"title"`
	Start  string `json:"start"`
	End    string `json:"end"`
	AllDay bool   `json:"allDay"`
	Color  string `json:"color"`
}

func (JadwalRapat) TableName() string {
	return "jadwal_rapats"
}

type JadwalCuti struct {
	ID     uint   `gorm:"primaryKey" json:"id"`
	Title  string `json:"title"`
	Start  string `json:"start"`
	End    string `json:"end"`
	AllDay bool   `json:"allDay"`
	Color  string `json:"color"` // Tambahkan field ini untuk warna
}

type TimelineDesktop struct {
	ID         uint   `gorm:"primaryKey" json:"id"`
	Start      string `json:"start"`
	End        string `json:"end"`
	ResourceId int    `json:"resourceId"` // Ubah tipe data dari string ke int
	Title      string `json:"title"`
	BgColor    string `json:"bgColor"`
}

func (TimelineDesktop) TableName() string {
	return "timeline_desktops"
}

type ResourceDesktop struct {
	ID       uint   `gorm:"primaryKey" json:"id"`
	Name     string `json:"name"`
	ParentID uint   `json:"parent_id"`
}

type File struct {
	ID          uint      `gorm:"primaryKey"`         // ID unik untuk file
	CreatedAt   time.Time `gorm:"autoCreateTime"`     // Timestamp saat file diunggah
	UpdatedAt   time.Time `gorm:"autoUpdateTime"`     // Timestamp untuk setiap update
	UserID      uint      `gorm:"index"`              // ID pengguna yang mengunggah file
	FilePath    string    `gorm:"not null"`           // Path lengkap di mana file disimpan
	FileName    string    `gorm:"not null"`           // Nama file asli
	ContentType string    `gorm:"not null"`           // Jenis konten file, misal 'application/pdf'
	Size        int64     `gorm:"not null"`           // Ukuran file dalam byte
}

// TableName overrides the table name used by File to `files`, if you want to specify it explicitly
func (File) TableName() string {
	return "files"
}

type Notification struct {
	ID       uint      `gorm:"primaryKey" json:"id"`
	Title    string    `json:"title"`
	Start    time.Time `json:"start"`
	Category string    `json:"category"`
}