package models

import (
	"encoding/json"
	"time"

	"gorm.io/gorm"
)

// model for suratMasuk
type SuratMasuk struct {
	ID         uint       `gorm:"primaryKey"`
	CreatedAt  *time.Time `gorm:"autoCreateTime"`
	UpdatedAt  *time.Time `gorm:"autoUpdateTime"`
	NoSurat    *string    `json:"no_surat"`
	Title      *string    `json:"title"`
	RelatedDiv *string    `json:"related_div"`
	DestinyDiv *string    `json:"destiny_div"`
	Tanggal    *time.Time `json:"tanggal"`
	CreateBy   string     `json:"create_by"`
}

func (i *SuratMasuk) MarshalJSON() ([]byte, error) {
	type Alias SuratMasuk
	if i.Tanggal == nil {
		// Handle jika Tanggal nil
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau format default yang diinginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

// model for suratKeluar
type SuratKeluar struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Title     *string    `json:"title"`
	From      *string    `json:"from"`
	Pic       *string    `json:"pic"`
	Tanggal   *time.Time `json:"tanggal"`
	CreateBy  string     `json:"create_by"`
}

func (i *SuratKeluar) MarshalJSON() ([]byte, error) {
	type Alias SuratKeluar
	tanggalFormatted := i.Tanggal.Format("2006-01-02")
	return json.Marshal(&struct {
		Tanggal *string `json:"tanggal"`
		*Alias
	}{
		Tanggal: &tanggalFormatted,
		Alias:   (*Alias)(i),
	})
}

type Arsip struct {
	gorm.Model
	NoArsip           *string    `json:"no_arsip"`
	JenisDokumen      *string    `json:"jenis_dokumen"`
	NoDokumen         *string    `json:"no_dokumen"`
	TanggalDokumen    *time.Time `json:"tanggal_dokumen"`
	Perihal           *string    `json:"perihal"`
	NoBox             *string    `json:"no_box"`
	TanggalPenyerahan *time.Time `json:"tanggal_penyerahan"`
	Keterangan        *string    `json:"keterangan"`
	CreateBy          string     `json:"create_by"`
}

func (a *Arsip) MarshalJSON() ([]byte, error) {
	type Alias Arsip
	var tanggalDokumenFormatted, tanggalPenyerahanFormatted string

	// Cek TanggalDokumen
	if a.TanggalDokumen == nil {
		tanggalDokumenFormatted = ""
	} else {
		tanggalDokumenFormatted = a.TanggalDokumen.Format("2006-01-02")
	}

	// Cek TanggalPenyerahan
	if a.TanggalPenyerahan == nil {
		tanggalPenyerahanFormatted = ""
	} else {
		tanggalPenyerahanFormatted = a.TanggalPenyerahan.Format("2006-01-02")
	}

	return json.Marshal(&struct {
		TanggalDokumen    string `json:"tanggal_dokumen"`
		TanggalPenyerahan string `json:"tanggal_penyerahan"`
		*Alias
	}{
		TanggalDokumen:    tanggalDokumenFormatted,
		TanggalPenyerahan: tanggalPenyerahanFormatted,
		Alias:             (*Alias)(a),
	})
}

type File struct {
	ID          uint      `gorm:"primaryKey"`     // ID unik untuk file
	CreatedAt   time.Time `gorm:"autoCreateTime"` // Timestamp saat file diunggah
	UpdatedAt   time.Time `gorm:"autoUpdateTime"` // Timestamp untuk setiap update
	UserID      uint      `gorm:"index"`          // ID pengguna yang mengunggah file
	FilePath    string    `gorm:"not null"`       // Path lengkap di mana file disimpan
	FileName    string    `gorm:"not null"`       // Nama file asli
	ContentType string    `gorm:"not null"`       // Jenis konten file, misal 'application/pdf'
	Size        int64     `gorm:"not null"`       // Ukuran file dalam byte
}

// TableName overrides the table name used by File to `files`, if you want to specify it explicitly
func (File) TableName() string {
	return "files"
}
