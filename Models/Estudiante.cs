using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace RegistroDeCompetenciaBash.Models
{
    public class Estudiante
    {
        //Keys
        [Key]
        [DataType(DataType.Text)]
        [Display(Name = "Número de estudiante")]
        [Required(ErrorMessage="Número de estudiante es requerido")]
        public string Id { get; set; }

        [ForeignKey("Recinto")]
        [DataType(DataType.Text)]
        [Display(Name = "Recinto")]
        [Required(ErrorMessage="Recinto es requerido")]
        public int? RecintoId { get; set; }

        //Attributes
        [Display(Name = "Nombre")]
        [DataType(DataType.Text)]
        [Required(ErrorMessage="Nombre es requerido")]
        public string Nombre { get; set; }

        [Display(Name = "Apellido paterno")]
        [DataType(DataType.Text)]
        [Required(ErrorMessage="Apellido paterno es requerido")]
        public string ApellidoPaterno { get; set; }

        [Display(Name = "Apellido materno")]
        [DataType(DataType.Text)]
        [Required(ErrorMessage="Apellido materno es requerido")]
        public string ApellidoMaterno { get; set; }

        [Display(Name = "Email")]
        [DataType(DataType.EmailAddress)]
        [Required(ErrorMessage="Email es requerido")]
        public string Email { get; set; }


        //Relations
        public virtual Recinto Recinto { get; set; } 
        
    }
}