using System;
using System.Runtime.Serialization;

namespace b2xtranslator.ZipUtils 
{

	/// <summary>Thrown whenever an error occurs during the build.</summary>
	[Serializable]
	public class ZipException : ApplicationException {

		/// <summary>Constructs an exception with no descriptive information.</summary>
		public ZipException() : base() {
		}

		/// <summary>Constructs an exception with a descriptive message.</summary>
		/// <param name="message">The error message that explains the reason for the exception.</param>
		public ZipException(string message) : base(message) {
		}

		/// <summary>Constructs an exception with a descriptive message and a reference to the instance of the <c>Exception</c> that is the root cause of the this exception.</summary>
		/// <param name="message">The error message that explains the reason for the exception.</param>
		/// <param name="innerException">An instance of <c>Exception</c> that is the cause of the current Exception. If <paramref name="innerException"/> is non-null, then the current Exception is raised in a catch block handling <paramref>innerException</paramref>.</param>
		public ZipException(string message, Exception innerException) : base(message, innerException) {
		}

		/// <summary>Initializes a new instance of the exception class with serialized data.</summary>
		/// <param name="info">The object that holds the serialized object data.</param>
		/// <param name="context">The contextual information about the source or destination.</param>
		public ZipException(SerializationInfo info, StreamingContext context) : base(info, context) {
		}
	}

    [Serializable]
    public class ZipCreationException : ZipException
    {
        /// <summary>Constructs an exception with no descriptive information.</summary>
        public ZipCreationException()
            : base()
        {
        }

        /// <summary>Constructs an exception with a descriptive message.</summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        public ZipCreationException(string message)
            : base(message)
        {
        }

        /// <summary>Constructs an exception with a descriptive message and a reference to the instance of the <c>Exception</c> that is the root cause of the this exception.</summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">An instance of <c>Exception</c> that is the cause of the current Exception. If <paramref name="innerException"/> is non-null, then the current Exception is raised in a catch block handling <paramref>innerException</paramref>.</param>
        public ZipCreationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>Initializes a new instance of the exception class with serialized data.</summary>
        /// <param name="info">The object that holds the serialized object data.</param>
        /// <param name="context">The contextual information about the source or destination.</param>
        public ZipCreationException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

    }
}
