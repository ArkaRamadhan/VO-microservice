package main

import (
	"project-its/common/middleware"
	"project-its/dokumen/internal/controllers"
	"project-its/dokumen/internal/initializers"
	"time"

	"github.com/gin-contrib/cors"
	"github.com/gin-contrib/sessions"
	"github.com/gin-contrib/sessions/cookie"
	"github.com/gin-gonic/gin"
)

func init() {

	initializers.LoadEnvVariables()
	initializers.ConnectToDB()

}

func main() {

	r := gin.Default()

	// Enable CORS
	r.Use(cors.New(cors.Config{
		AllowOrigins:     []string{"http://localhost:8000"},
		AllowMethods:     []string{"GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"},
		AllowHeaders:     []string{"Origin", "Content-Type", "Accept", "Authorization"},
		ExposeHeaders:    []string{"Content-Length"},
		AllowCredentials: true,
		AllowOriginFunc: func(origin string) bool {
			return origin == "http://localhost:8000"
		},
		MaxAge: 12 * time.Hour,
	}))

	r.Use(middleware.TokenAuthMiddleware())

	// Routes for User
	store := cookie.NewStore([]byte("secret"))
	r.Use(sessions.Sessions("mysession", store))

	// ****************** Memo ******************//
	r.GET("/memo", controllers.MemoIndex)
	r.POST("/memo", controllers.MemoCreate)
	r.GET("/memo/:id", controllers.MemoShow)
	r.PUT("/memo/:id", controllers.MemoUpdate)
	r.DELETE("/memo/:id", controllers.MemoDelete)
	r.GET("/exportMemo", controllers.ExportMemoHandler)
	r.POST("/uploadMemo", controllers.ImportExcelMemo)
	r.POST("/uploadFileMemo", controllers.UploadHandlerMemo)
	r.GET("/downloadMemo/:id/:filename", controllers.DownloadFileHandlerMemo)
	r.DELETE("/deleteMemo/:id/:filename", controllers.DeleteFileHandlerMemo)
	r.GET("/filesMemo/:id", controllers.GetFilesByIDMemo)
	// ****************** End Memo ******************//

	// ****************** Berita Acara ******************//
	r.GET("/beritaAcara", controllers.BeritaAcaraIndex)
	r.POST("/beritaAcara", controllers.BeritaAcaraCreate)
	r.GET("/beritaAcara/:id", controllers.BeritaAcaraShow)
	r.PUT("/beritaAcara/:id", controllers.BeritaAcaraUpdate)
	r.DELETE("/beritaAcara/:id", controllers.BeritaAcaraDelete)
	r.GET("/exportBeritaAcara", controllers.ExportBeritaAcaraHandler)
	r.POST("/uploadBeritaAcara", controllers.ImportExcelBeritaAcara)
	r.POST("/uploadFileBeritaAcara", controllers.UploadHandlerBeritaAcara)
	r.GET("/downloadBeritaAcara/:id/:filename", controllers.DownloadFileHandlerBeritaAcara)
	r.DELETE("/deleteBeritaAcara/:id/:filename", controllers.DeleteFileHandlerBeritaAcara)
	r.GET("/filesBeritaAcara/:id", controllers.GetFilesByIDBeritaAcara)
	// ****************** End Berita Acara ******************//

	// ****************** Surat ******************//
	r.GET("/surat", controllers.SuratIndex)
	r.POST("/surat", controllers.SuratCreate)
	r.GET("/surat/:id", controllers.SuratShow)
	r.PUT("/surat/:id", controllers.SuratUpdate)
	r.DELETE("/surat/:id", controllers.SuratDelete)
	r.GET("/exportSurat", controllers.ExportSuratHandler)
	r.POST("/uploadSurat", controllers.ImportExcelSurat)
	r.POST("/uploadFileSurat", controllers.UploadHandlerSurat)
	r.GET("/downloadSurat/:id/:filename", controllers.DownloadFileHandlerSurat)
	r.DELETE("/deleteSurat/:id/:filename", controllers.DeleteFileHandlerSurat)
	r.GET("/filesSurat/:id", controllers.GetFilesByIDSurat)
	// ****************** End Surat ******************//

	// ****************** SK ******************//
	r.GET("/sk", controllers.SkIndex)
	r.POST("/sk", controllers.SkCreate)
	r.GET("/sk/:id", controllers.SkShow)
	r.PUT("/sk/:id", controllers.SkUpdate)
	r.DELETE("/sk/:id", controllers.SkDelete)
	r.GET("/exportSk", controllers.ExportSkHandler)
	r.POST("/uploadSk", controllers.ImportExcelSk)
	r.POST("/uploadFileSk", controllers.UploadHandlerSk)
	r.GET("/downloadSk/:id/:filename", controllers.DownloadFileHandlerSk)
	r.DELETE("/deleteSk/:id/:filename", controllers.DeleteFileHandlerSk)
	r.GET("/filesSk/:id", controllers.GetFilesByIDSk)
	// ****************** End SK ******************//
}
